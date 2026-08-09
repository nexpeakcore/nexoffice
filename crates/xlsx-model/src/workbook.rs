//! sparse workbook containers and the calc-facing cell-access trait.

use std::collections::{BTreeMap, BTreeSet};
use std::ops::Range;

use serde::{Deserialize, Serialize};

use crate::addr::{CellRange, CellRef, ColId, RowId, SheetId};
use crate::date::DateSystem;
use crate::styles::Stylesheet;
use crate::value::CellValue;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct FreezePane {
    pub rows: RowId,
    pub cols: ColId,
    pub top_left: CellRef,
}

impl FreezePane {
    pub fn new(rows: RowId, cols: ColId, top_left: CellRef) -> Self {
        Self {
            rows,
            cols,
            top_left,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct DefinedName {
    pub name: String,
    pub formula: String,
    pub local_sheet: Option<SheetId>,
    pub hidden: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct AutoFilter {
    pub range: CellRange,
    /// per-column criteria; empty when the filter only adds dropdowns.
    pub columns: Vec<AutoFilterColumn>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct AutoFilterColumn {
    /// absolute sheet column; the xml `colId` is relative to `range.start`.
    pub col: ColId,
    /// explicit cell texts kept visible. `None` places no constraint on a row;
    /// `Some(list)` is an allow-list, and an empty list is spreadsheetml's
    /// blanks-only filter, which only keeps rows when `show_blanks`.
    pub values: Option<Vec<String>>,
    /// whether blank cells stay visible alongside `values`.
    pub show_blanks: bool,
    /// criteria this engine does not model — `customFilters`, `top10`,
    /// `dynamicFilter`, `colorFilter`, `iconFilter`, date-group items — kept as
    /// the source `filterColumn`'s inner xml so a save writes them back
    /// unchanged. mutually exclusive with `values`: such a column constrains
    /// nothing, so re-evaluating the filter never hides a row on its account.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unsupported: Option<String>,
}

impl AutoFilterColumn {
    /// the allow-list the column narrows rows to, or `None` when it constrains
    /// nothing: either it carries no criteria or its criteria are only
    /// preserved (`unsupported`) rather than evaluated.
    pub fn criteria(&self) -> Option<&[String]> {
        self.values.as_deref()
    }
}

/// a classic cell note: plain text plus its author. rich runs collapse to
/// concatenated text.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Comment {
    pub author: String,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Hyperlink {
    pub range: CellRange,
    pub external_target: Option<String>,
    pub location: Option<String>,
    pub tooltip: Option<String>,
    pub display: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Cell {
    pub value: CellValue,
    /// original formula text without the leading `=`, if any.
    pub formula: Option<String>,
    /// index into the workbook style table (cellXfs).
    pub style: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Sheet {
    pub name: String,
    cells: BTreeMap<(RowId, ColId), Cell>,
    pub freeze_pane: Option<FreezePane>,
    pub hyperlinks: Vec<Hyperlink>,
    pub merges: Vec<CellRange>,
    pub col_widths: BTreeMap<ColId, f64>,
    pub row_heights: BTreeMap<RowId, f64>,
    /// rows explicitly hidden (by a filter or a manual hide).
    pub hidden_rows: BTreeSet<RowId>,
    /// the subset of `hidden_rows` the user hid by hand. SpreadsheetML records
    /// only `hidden="1"` per row, so provenance cannot be serialized: this is
    /// seeded at load from the hidden rows that pass the active filter and is
    /// then maintained exactly for as long as the workbook stays open, which
    /// is what keeps a manual hide hidden when the filter changes or clears.
    pub manual_hidden_rows: BTreeSet<RowId>,
    pub auto_filter: Option<AutoFilter>,
    pub comments: BTreeMap<(RowId, ColId), Comment>,
    /// Charts anchored on this sheet, parsed for display only — edits never
    /// touch them and saves re-emit the preserved source parts.
    pub drawings: Vec<SheetDrawing>,
}

/// One cell-anchored drawing object (today: only charts are modeled).
#[derive(Clone, Debug, PartialEq)]
pub struct SheetDrawing {
    pub anchor: DrawingAnchor,
    pub chart: ooxml_drawingml::chart::ChartSpace,
    /// `false` for drawings parsed from a preserved package (their source
    /// parts ride through saves); `true` for drawings authored in-session,
    /// which saves must serialize into new parts.
    pub created: bool,
}

/// Where a drawing sits on the grid, in cells plus EMU offsets.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DrawingAnchor {
    /// `twoCellAnchor`/`oneCellAnchor`: pinned to a grid cell, sized by a
    /// second cell or an explicit `<xdr:ext cx cy>`.
    Cell {
        from: AnchorCell,
        to: Option<AnchorCell>,
        extent_emu: Option<(i64, i64)>,
    },
    /// `absoluteAnchor`: a fixed sheet position and size in EMU.
    Absolute {
        pos_emu: (i64, i64),
        extent_emu: (i64, i64),
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct AnchorCell {
    pub col: u32,
    pub col_offset_emu: i64,
    pub row: u32,
    pub row_offset_emu: i64,
}

impl Sheet {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            ..Self::default()
        }
    }

    pub fn cell(&self, at: CellRef) -> Option<&Cell> {
        self.cells.get(&(at.row, at.col))
    }

    pub fn set_cell(&mut self, at: CellRef, cell: Cell) {
        if cell == Cell::default() {
            self.cells.remove(&(at.row, at.col));
        } else {
            self.cells.insert((at.row, at.col), cell);
        }
    }

    /// ordered iteration over occupied cells (row-major).
    pub fn iter_cells(&self) -> impl Iterator<Item = (CellRef, &Cell)> {
        self.cells
            .iter()
            .map(|(&(row, col), cell)| (CellRef::new(row, col), cell))
    }

    pub fn iter_cells_in_rect(
        &self,
        rows: Range<RowId>,
        cols: Range<ColId>,
    ) -> impl Iterator<Item = (CellRef, &Cell)> {
        let start_col = cols.start;
        let end_col = cols.end.max(start_col);
        rows.flat_map(move |row| {
            self.cells
                .range((row, start_col)..(row, end_col))
                .map(|(&(row, col), cell)| (CellRef::new(row, col), cell))
        })
    }

    /// hide `row` by hand: it stays hidden across filter changes until it is
    /// shown again.
    pub fn hide_row(&mut self, row: RowId) {
        self.hidden_rows.insert(row);
        self.manual_hidden_rows.insert(row);
    }

    /// show `row` by hand, dropping any manual hide on it. a row the active
    /// filter rejects re-hides the next time the filter is evaluated.
    pub fn show_row(&mut self, row: RowId) {
        self.hidden_rows.remove(&row);
        self.manual_hidden_rows.remove(&row);
    }

    pub fn is_row_hidden(&self, row: RowId) -> bool {
        self.hidden_rows.contains(&row)
    }

    /// the non-header rows of `filter`'s range that fail its criteria against
    /// this sheet's current cells — the rows the filter itself hides. the
    /// filter is passed in because callers evaluate a candidate filter before
    /// installing it.
    pub fn rows_failing_filter(&self, filter: Option<&AutoFilter>) -> BTreeSet<RowId> {
        let Some(filter) = filter else {
            return BTreeSet::new();
        };
        (filter.range.start.row..=filter.range.end.row)
            .filter(|&row| row != filter.range.start.row && !self.row_passes_filter(filter, row))
            .collect()
    }

    /// a row is visible when every column with value criteria matches its cell
    /// text, blanks counting only when the column keeps blanks. a column that
    /// constrains nothing — no criteria at all, or criteria this engine only
    /// preserves rather than evaluates — always passes.
    fn row_passes_filter(&self, filter: &AutoFilter, row: RowId) -> bool {
        filter.columns.iter().all(|column| {
            let Some(values) = column.criteria() else {
                return true;
            };
            let text = self
                .cell(CellRef::new(row, column.col))
                .map(|cell| filter_text(&cell.value))
                .unwrap_or_default();
            if text.is_empty() {
                column.show_blanks
            } else {
                values.contains(&text)
            }
        })
    }

    /// recover manual-hide provenance for a freshly loaded sheet. the file
    /// says only that a row is hidden, so a hidden row the active filter would
    /// keep visible must have been hidden by hand; a hidden row the filter
    /// rejects is indistinguishable from a filter hide and is attributed to
    /// the filter.
    pub fn seed_manual_hidden_rows(&mut self) {
        let failing = self.rows_failing_filter(self.auto_filter.as_ref());
        self.manual_hidden_rows = self.hidden_rows.difference(&failing).copied().collect();
    }

    pub fn comment_at(&self, at: CellRef) -> Option<&Comment> {
        self.comments.get(&(at.row, at.col))
    }

    /// set, replace, or (`None`) delete the comment at `at`, returning the
    /// previous one.
    pub fn set_comment(&mut self, at: CellRef, comment: Option<Comment>) -> Option<Comment> {
        match comment {
            Some(comment) => self.comments.insert((at.row, at.col), comment),
            None => self.comments.remove(&(at.row, at.col)),
        }
    }

    pub fn hyperlink_at(&self, at: CellRef) -> Option<&Hyperlink> {
        self.hyperlinks.iter().find(|link| link.range.contains(at))
    }

    /// the rectangle covering everything anchored to a cell: values,
    /// hyperlinks, and comments. a note on an otherwise empty cell counts, so
    /// the grid a viewer scrolls always reaches far enough to show it.
    pub fn used_range(&self) -> Option<CellRange> {
        let mut bounds = self
            .cells
            .keys()
            .chain(self.comments.keys())
            .map(|&(row, col)| CellRange::new(CellRef::new(row, col), CellRef::new(row, col)))
            .chain(self.hyperlinks.iter().map(|link| link.range));
        let first = bounds.next()?;
        let (mut min_r, mut max_r, mut min_c, mut max_c) = (
            first.start.row,
            first.end.row,
            first.start.col,
            first.end.col,
        );
        for range in bounds {
            let r = range.start.row;
            let c = range.start.col;
            min_r = min_r.min(r);
            max_r = max_r.max(range.end.row);
            min_c = min_c.min(c);
            max_c = max_c.max(range.end.col);
        }
        Some(CellRange::new(
            CellRef::new(min_r, min_c),
            CellRef::new(max_r, max_c),
        ))
    }
}

/// the text an auto filter matches against: the bare value, no number formats.
pub fn filter_text(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number { value } => value.to_string(),
        CellValue::Text { value } => value.clone(),
        CellValue::Bool { value } => if *value { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::Error { value } => value.as_str().to_string(),
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub date_system: DateSystem,
    pub defined_names: Vec<DefinedName>,
    /// shared string table as parsed; kept for round-trip fidelity.
    pub shared_strings: Vec<String>,
    /// parsed style tables + theme; a cell's `style` indexes `styles.cell_xfs`.
    pub styles: Stylesheet,
}

impl Workbook {
    pub fn sheet(&self, id: SheetId) -> Option<&Sheet> {
        self.sheets.get(id.0 as usize)
    }

    pub fn sheet_mut(&mut self, id: SheetId) -> Option<&mut Sheet> {
        self.sheets.get_mut(id.0 as usize)
    }

    pub fn sheet_by_name(&self, name: &str) -> Option<(SheetId, &Sheet)> {
        let name = name.to_lowercase();
        self.sheets
            .iter()
            .enumerate()
            .find(|(_, sheet)| sheet.name.to_lowercase() == name)
            .map(|(i, s)| (SheetId(i as u32), s))
    }

    pub fn defined_name(&self, sheet: SheetId, name: &str) -> Option<&DefinedName> {
        self.defined_names
            .iter()
            .find(|defined| {
                defined.local_sheet == Some(sheet) && defined.name.eq_ignore_ascii_case(name)
            })
            .or_else(|| {
                self.defined_names.iter().find(|defined| {
                    defined.local_sheet.is_none() && defined.name.eq_ignore_ascii_case(name)
                })
            })
    }
}

/// read access the calc engine evaluates through.
pub trait CellProvider {
    fn value(&self, sheet: SheetId, at: CellRef) -> CellValue;
    fn formula(&self, sheet: SheetId, at: CellRef) -> Option<&str>;
    fn sheet_id(&self, name: &str) -> Option<SheetId>;
    fn defined_name(&self, _sheet: SheetId, _name: &str) -> Option<&DefinedName> {
        None
    }
}

impl CellProvider for Workbook {
    fn value(&self, sheet: SheetId, at: CellRef) -> CellValue {
        self.sheet(sheet)
            .and_then(|s| s.cell(at))
            .map(|c| c.value.clone())
            .unwrap_or_default()
    }

    fn formula(&self, sheet: SheetId, at: CellRef) -> Option<&str> {
        self.sheet(sheet)?.cell(at)?.formula.as_deref()
    }

    fn sheet_id(&self, name: &str) -> Option<SheetId> {
        self.sheet_by_name(name).map(|(id, _)| id)
    }

    fn defined_name(&self, sheet: SheetId, name: &str) -> Option<&DefinedName> {
        self.defined_name(sheet, name)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sparse_set_get_and_used_range() {
        let mut sheet = Sheet::new("Sheet1");
        assert!(sheet.used_range().is_none());

        let b2 = CellRef::parse_a1("B2").unwrap();
        let d7 = CellRef::parse_a1("D7").unwrap();
        sheet.set_cell(
            b2,
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
        sheet.set_cell(
            d7,
            Cell {
                value: CellValue::Text { value: "x".into() },
                ..Cell::default()
            },
        );

        assert_eq!(sheet.used_range().unwrap().to_a1(), "B2:D7");
        assert_eq!(sheet.iter_cells().count(), 2);

        sheet.set_cell(b2, Cell::default());
        assert_eq!(sheet.used_range().unwrap().to_a1(), "D7");
    }

    #[test]
    fn workbook_cell_provider() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.defined_names.push(DefinedName {
            name: "Answer".into(),
            formula: "A1".into(),
            local_sheet: None,
            hidden: false,
        });
        let a1 = CellRef::parse_a1("A1").unwrap();
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            a1,
            Cell {
                value: CellValue::Number { value: 42.0 },
                formula: Some("40+2".into()),
                style: None,
            },
        );

        let id = wb.sheet_id("Data").unwrap();
        assert_eq!(wb.sheet_id("data"), Some(id));
        assert_eq!(wb.value(id, a1), CellValue::Number { value: 42.0 });
        assert_eq!(wb.formula(id, a1), Some("40+2"));
        assert_eq!(
            wb.value(id, CellRef::parse_a1("Z9").unwrap()),
            CellValue::Empty
        );
        assert!(wb.sheet_id("Nope").is_none());
        assert_eq!(
            CellProvider::defined_name(&wb, id, "answer").map(|defined| defined.formula.as_str()),
            Some("A1")
        );
    }

    #[test]
    fn local_defined_name_shadows_workbook_name() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.defined_names.extend([
            DefinedName {
                name: "Rate".into(),
                formula: "1".into(),
                local_sheet: None,
                hidden: false,
            },
            DefinedName {
                name: "rate".into(),
                formula: "2".into(),
                local_sheet: Some(SheetId(0)),
                hidden: false,
            },
        ]);

        assert_eq!(
            wb.defined_name(SheetId(0), "RATE")
                .map(|defined| defined.formula.as_str()),
            Some("2")
        );
        assert_eq!(
            wb.defined_name(SheetId(1), "RATE")
                .map(|defined| defined.formula.as_str()),
            Some("1")
        );
    }

    #[test]
    fn iterates_only_cells_in_rectangle() {
        let mut sheet = Sheet::new("Data");
        for address in ["A1", "B2", "C3", "Z100"] {
            sheet.set_cell(
                CellRef::parse_a1(address).unwrap(),
                Cell {
                    value: CellValue::Number { value: 1.0 },
                    ..Cell::default()
                },
            );
        }
        let cells: Vec<_> = sheet
            .iter_cells_in_rect(0..3, 0..2)
            .map(|(cell, _)| cell.to_a1())
            .collect();
        assert_eq!(cells, vec!["A1", "B2"]);
        let mut reversed = 1..2;
        std::mem::swap(&mut reversed.start, &mut reversed.end);
        assert_eq!(sheet.iter_cells_in_rect(0..3, reversed).count(), 0);
    }

    #[test]
    fn hides_and_shows_rows() {
        let mut sheet = Sheet::new("Data");
        assert!(!sheet.is_row_hidden(3));

        sheet.hide_row(3);
        sheet.hide_row(5);
        assert!(sheet.is_row_hidden(3));
        assert!(sheet.is_row_hidden(5));
        assert!(!sheet.is_row_hidden(4));

        sheet.show_row(3);
        assert!(!sheet.is_row_hidden(3));
        assert!(sheet.is_row_hidden(5));
    }

    #[test]
    fn comments_set_replace_and_delete() {
        let mut sheet = Sheet::new("Data");
        let b2 = CellRef::parse_a1("B2").unwrap();
        assert!(sheet.comment_at(b2).is_none());

        let first = Comment {
            author: "Ada".into(),
            text: "check this".into(),
        };
        assert_eq!(sheet.set_comment(b2, Some(first.clone())), None);
        assert_eq!(sheet.comment_at(b2), Some(&first));

        let second = Comment {
            author: "Grace".into(),
            text: "done".into(),
        };
        assert_eq!(sheet.set_comment(b2, Some(second.clone())), Some(first));
        assert_eq!(sheet.set_comment(b2, None), Some(second));
        assert!(sheet.comments.is_empty());
    }

    #[test]
    fn hyperlinks_are_addressable_and_extend_the_used_range() {
        let mut sheet = Sheet::new("Data");
        sheet.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("C4:D5").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: Some("Example".into()),
        });

        assert_eq!(sheet.used_range().unwrap().to_a1(), "C4:D5");
        assert_eq!(
            sheet
                .hyperlink_at(CellRef::parse_a1("D5").unwrap())
                .and_then(|link| link.display.as_deref()),
            Some("Example")
        );
        assert!(
            sheet
                .hyperlink_at(CellRef::parse_a1("A1").unwrap())
                .is_none()
        );
    }

    /// A note on an otherwise empty cell has to sit inside the used range, or
    /// the editor never lays out a rect for its indicator and the note becomes
    /// invisible and uneditable.
    #[test]
    fn comments_extend_the_used_range() {
        let mut sheet = Sheet::new("Data");
        sheet.set_cell(
            CellRef::parse_a1("B2").unwrap(),
            Cell {
                value: CellValue::Text {
                    value: "only".into(),
                },
                ..Cell::default()
            },
        );
        sheet.set_comment(
            CellRef::parse_a1("E9").unwrap(),
            Some(Comment {
                author: "Ada".into(),
                text: "look here".into(),
            }),
        );

        assert_eq!(sheet.used_range().unwrap().to_a1(), "B2:E9");
    }

    #[test]
    fn a_comment_alone_gives_an_otherwise_empty_sheet_an_extent() {
        let mut sheet = Sheet::new("Data");
        sheet.set_comment(
            CellRef::parse_a1("C3").unwrap(),
            Some(Comment {
                author: "Ada".into(),
                text: "note".into(),
            }),
        );

        assert_eq!(sheet.used_range().unwrap().to_a1(), "C3");
    }

    /// Criteria the engine only preserves must not narrow anything: an empty
    /// allow-list would hide every non-blank row in the filter's range.
    #[test]
    fn unsupported_filter_criteria_hide_no_rows() {
        let mut sheet = Sheet::new("Data");
        for row in 0..4 {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number {
                        value: f64::from(row),
                    },
                    ..Cell::default()
                },
            );
        }
        let column = AutoFilterColumn {
            col: 0,
            values: None,
            show_blanks: true,
            unsupported: Some(r#"<top10 top="1" val="2"/>"#.into()),
        };
        assert!(column.criteria().is_none());
        sheet.auto_filter = Some(AutoFilter {
            range: CellRange::parse_a1("A1:A4").unwrap(),
            columns: vec![column],
        });

        assert!(
            sheet
                .rows_failing_filter(sheet.auto_filter.as_ref())
                .is_empty()
        );
    }

    #[test]
    fn manual_hide_provenance_is_seeded_from_the_filter_and_tracked_by_hand() {
        let mut sheet = Sheet::new("Data");
        for (row, name) in [(0, "Name"), (1, "keep"), (2, "drop"), (3, "keep")] {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Text { value: name.into() },
                    ..Cell::default()
                },
            );
        }
        sheet.auto_filter = Some(AutoFilter {
            range: CellRange::parse_a1("A1:A4").unwrap(),
            columns: vec![AutoFilterColumn {
                col: 0,
                values: Some(vec!["keep".into()]),
                show_blanks: false,
                unsupported: None,
            }],
        });
        assert_eq!(
            sheet.rows_failing_filter(sheet.auto_filter.as_ref()),
            [2].into_iter().collect::<BTreeSet<_>>(),
            "the header never fails and only the `drop` row does"
        );

        sheet.hidden_rows = [2, 3].into_iter().collect();
        sheet.seed_manual_hidden_rows();
        assert_eq!(
            sheet.manual_hidden_rows,
            [3].into_iter().collect::<BTreeSet<_>>(),
            "a hidden row the filter would keep must have been hidden by hand"
        );

        sheet.hide_row(1);
        assert!(sheet.is_row_hidden(1));
        assert!(sheet.manual_hidden_rows.contains(&1));
        sheet.show_row(1);
        assert!(!sheet.is_row_hidden(1));
        assert!(!sheet.manual_hidden_rows.contains(&1));
    }
}
