//! applying ops mutably to a workbook, returning the inverse for undo, plus
//! `remap_ref` — the shared address-remapping primitive.

use std::cmp::Ordering;
use std::collections::{BTreeMap, BTreeSet};
use std::fmt;

use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{
    AutoFilter, Cell, CellRange, CellRef, CellValue, ColId, Comment, DefinedName, RowId, Sheet,
    SheetId, Workbook,
};

use crate::formatting::{mutate_number_format, patch_cell_format};
use crate::op::{CellState, Op};
use crate::remap::{
    remap_defined_names, remap_formulas, remap_hyperlink_locations, remap_hyperlink_range,
    rename_defined_names, rename_hyperlink_location, rename_sheet_references, shift_formula_rows,
};

/// the inverse of an applied op: a base-vocabulary op list that, replayed in
/// order, restores the prior workbook state.
#[derive(Debug, Clone, PartialEq)]
pub struct InvertedOp(pub Vec<Op>);

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OpError {
    SheetNotFound(SheetId),
    SheetIndexOutOfRange(usize),
    FormulaNotRewritable {
        sheet: SheetId,
        cell: CellRef,
    },
    DefinedNameNotRewritable {
        name: String,
    },
    InvalidStyle(String),
    MergeConflict {
        sheet: SheetId,
        merge: CellRange,
    },
    HyperlinkConflict {
        sheet: SheetId,
        hyperlink: CellRange,
    },
}

impl fmt::Display for OpError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            OpError::SheetNotFound(id) => write!(f, "sheet {} not found", id.0),
            OpError::SheetIndexOutOfRange(i) => write!(f, "sheet index {i} out of range"),
            OpError::FormulaNotRewritable { sheet, cell } => write!(
                f,
                "formula at sheet {}, {} cannot be safely rewritten",
                sheet.0,
                cell.to_a1()
            ),
            OpError::DefinedNameNotRewritable { name } => {
                write!(f, "defined name {name} cannot be safely rewritten")
            }
            OpError::InvalidStyle(message) => f.write_str(message),
            OpError::MergeConflict { sheet, merge } => write!(
                f,
                "merged range {} on sheet {} conflicts with the sort range; merges must lie \
                 inside it, span a single row, and repeat on every sorted row",
                merge.to_a1(),
                sheet.0
            ),
            OpError::HyperlinkConflict { sheet, hyperlink } => write!(
                f,
                "hyperlink range {} on sheet {} conflicts with the sort range; a hyperlink only \
                 moves with its row when it spans a single row inside the sorted columns",
                hyperlink.to_a1(),
                sheet.0
            ),
        }
    }
}

impl std::error::Error for OpError {}

/// apply one op, mutating `wb` and returning its inverse.
pub fn apply(wb: &mut Workbook, op: &Op) -> Result<InvertedOp, OpError> {
    match op {
        Op::SetCell { sheet, at, cell } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.cell(*at).map(CellState::from).unwrap_or_default();
            s.set_cell(*at, cell.clone().into());
            Ok(InvertedOp(vec![Op::SetCell {
                sheet: *sheet,
                at: *at,
                cell: old,
            }]))
        }
        Op::SetColWidth { sheet, col, width } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.col_widths.get(col).copied();
            match width {
                Some(w) => {
                    s.col_widths.insert(*col, *w);
                }
                None => {
                    s.col_widths.remove(col);
                }
            }
            Ok(InvertedOp(vec![Op::SetColWidth {
                sheet: *sheet,
                col: *col,
                width: old,
            }]))
        }
        Op::SetRowHeight { sheet, row, height } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.row_heights.get(row).copied();
            match height {
                Some(h) => {
                    s.row_heights.insert(*row, *h);
                }
                None => {
                    s.row_heights.remove(row);
                }
            }
            Ok(InvertedOp(vec![Op::SetRowHeight {
                sheet: *sheet,
                row: *row,
                height: old,
            }]))
        }
        Op::SetFreezePane { sheet, pane } => {
            let sheet_ref = sheet_mut(wb, *sheet)?;
            let old = sheet_ref.freeze_pane;
            sheet_ref.freeze_pane = *pane;
            Ok(InvertedOp(vec![Op::SetFreezePane {
                sheet: *sheet,
                pane: old,
            }]))
        }
        Op::SortRange {
            sheet,
            range,
            key_col,
            ascending,
        } => {
            let s = sheet_mut(wb, *sheet)?;
            if let Some(merge) = find_sort_merge_conflict(s, *range) {
                return Err(OpError::MergeConflict {
                    sheet: *sheet,
                    merge,
                });
            }
            let original: Vec<RowId> = (range.start.row..=range.end.row).collect();
            let mut order = original.clone();
            order.sort_by(|&a, &b| {
                compare_sort_keys(
                    s.cell(CellRef::new(a, *key_col)).map(|cell| &cell.value),
                    s.cell(CellRef::new(b, *key_col)).map(|cell| &cell.value),
                    *ascending,
                )
            });
            let moves: Vec<(RowId, RowId)> = order
                .into_iter()
                .zip(original)
                .filter(|(from, to)| from != to)
                .collect();
            let outcome = move_row_cells(s, *sheet, *range, &moves)?;
            Ok(InvertedOp(outcome.into_inverse(*sheet, *range)))
        }
        Op::MoveRows {
            sheet,
            range,
            moves,
        } => {
            let s = sheet_mut(wb, *sheet)?;
            let outcome = move_row_cells(s, *sheet, *range, moves)?;
            Ok(InvertedOp(outcome.into_inverse(*sheet, *range)))
        }
        Op::SetAutoFilter { sheet, filter } => {
            let s = sheet_mut(wb, *sheet)?;
            let old_filter = s.auto_filter.clone();
            let old_hidden: Vec<RowId> = s.hidden_rows.iter().copied().collect();
            let previously_filter_hidden = filter_failing_rows(s, old_filter.as_ref());
            s.auto_filter = filter.clone();
            let mut hidden: BTreeSet<RowId> = s
                .hidden_rows
                .iter()
                .copied()
                .filter(|row| !previously_filter_hidden.contains(row))
                .collect();
            hidden.extend(filter_failing_rows(s, filter.as_ref()));
            s.hidden_rows = hidden;
            Ok(InvertedOp(vec![
                Op::SetAutoFilter {
                    sheet: *sheet,
                    filter: old_filter,
                },
                Op::SetHiddenRows {
                    sheet: *sheet,
                    hidden: old_hidden,
                },
            ]))
        }
        Op::SetHiddenRows { sheet, hidden } => {
            let s = sheet_mut(wb, *sheet)?;
            let old: Vec<RowId> = s.hidden_rows.iter().copied().collect();
            s.hidden_rows = hidden.iter().copied().collect();
            Ok(InvertedOp(vec![Op::SetHiddenRows {
                sheet: *sheet,
                hidden: old,
            }]))
        }
        Op::SetComment {
            sheet,
            cell,
            comment,
        } => {
            let s = sheet_mut(wb, *sheet)?;
            let old = s.set_comment(*cell, comment.clone());
            Ok(InvertedOp(vec![Op::SetComment {
                sheet: *sheet,
                cell: *cell,
                comment: old,
            }]))
        }
        Op::SetHyperlinks { sheet, hyperlinks } => {
            let sheet_ref = sheet_mut(wb, *sheet)?;
            let old = std::mem::replace(&mut sheet_ref.hyperlinks, hyperlinks.clone());
            Ok(InvertedOp(vec![Op::SetHyperlinks {
                sheet: *sheet,
                hyperlinks: old,
            }]))
        }
        Op::MergeCells { sheet, range } => {
            let s = sheet_mut(wb, *sheet)?;
            let replaced = s
                .merges
                .iter()
                .copied()
                .filter(|merged| ranges_intersect(*merged, *range))
                .collect::<Vec<_>>();
            if replaced.as_slice() == [*range] {
                return Ok(InvertedOp(vec![]));
            }
            s.merges.retain(|merged| !ranges_intersect(*merged, *range));
            s.merges.push(*range);
            let mut inverse = Vec::with_capacity(replaced.len() + 1);
            inverse.push(Op::UnmergeCells {
                sheet: *sheet,
                range: *range,
            });
            inverse.extend(replaced.into_iter().map(|range| Op::MergeCells {
                sheet: *sheet,
                range,
            }));
            Ok(InvertedOp(inverse))
        }
        Op::UnmergeCells { sheet, range } => {
            let s = sheet_mut(wb, *sheet)?;
            match s.merges.iter().position(|m| m == range) {
                Some(pos) => {
                    s.merges.remove(pos);
                    Ok(InvertedOp(vec![Op::MergeCells {
                        sheet: *sheet,
                        range: *range,
                    }]))
                }
                None => Ok(InvertedOp(vec![])),
            }
        }
        Op::PatchRangeStyle {
            sheet,
            range,
            patch,
        } => apply_range_formats(wb, *sheet, *range, |format, row, col| {
            patch_cell_format(format, patch, *range, row, col)
        }),
        Op::SetRangeNumberFormat {
            sheet,
            range,
            format,
        } => apply_range_formats(wb, *sheet, *range, |cell_format, _, _| {
            mutate_number_format(cell_format, format);
            Ok(())
        }),
        Op::ApplyRangeFormat {
            sheet,
            range,
            format,
        } => {
            if format.rows == 0
                || format.columns == 0
                || format.formats.len() != (format.rows as usize) * (format.columns as usize)
            {
                return Err(OpError::InvalidStyle(
                    "captured format dimensions do not match its cells".into(),
                ));
            }
            apply_range_formats(wb, *sheet, *range, |cell_format, row, col| {
                let source_row = (row - range.start.row) % format.rows;
                let source_col = (col - range.start.col) % format.columns;
                let index = (source_row * format.columns + source_col) as usize;
                *cell_format = format.formats[index].clone();
                Ok(())
            })
        }
        Op::AddSheet { index, name } => {
            let idx = (*index).min(wb.sheets.len());
            wb.sheets.insert(idx, Sheet::new(name.clone()));
            shift_defined_name_scopes(wb, idx);
            Ok(InvertedOp(vec![Op::RemoveSheet { index: idx }]))
        }
        Op::RemoveSheet { index } => remove_sheet(wb, *index),
        Op::RenameSheet { sheet, name } => {
            let old = wb
                .sheet(*sheet)
                .ok_or(OpError::SheetNotFound(*sheet))?
                .name
                .clone();
            let formulas = rename_sheet_references(wb, &old, name)?;
            let hyperlinks = rename_hyperlink_locations(wb, &old, name);
            let defined_names = rename_defined_names(wb, &old, name);
            sheet_mut(wb, *sheet)?.name = name.clone();
            let mut inverse = vec![Op::RestoreSheet {
                sheet: *sheet,
                name: old,
                formulas,
            }];
            inverse.extend(hyperlinks);
            inverse.extend(defined_names);
            Ok(InvertedOp(inverse))
        }
        Op::RestoreSheet {
            sheet,
            name,
            formulas,
        } => {
            let old_name = wb
                .sheet(*sheet)
                .ok_or(OpError::SheetNotFound(*sheet))?
                .name
                .clone();
            let mut old_formulas = Vec::with_capacity(formulas.len());
            for (formula_sheet, cell, _) in formulas {
                let sheet_ref = wb
                    .sheet(*formula_sheet)
                    .ok_or(OpError::SheetNotFound(*formula_sheet))?;
                let state = sheet_ref
                    .cell(*cell)
                    .map(CellState::from)
                    .unwrap_or_default();
                old_formulas.push((*formula_sheet, *cell, state));
            }
            let defined_names = rename_defined_names(wb, &old_name, name);
            sheet_mut(wb, *sheet)?.name = name.clone();
            for (formula_sheet, cell, state) in formulas {
                sheet_mut(wb, *formula_sheet)?.set_cell(*cell, state.clone().into());
            }
            let mut inverse = vec![Op::RestoreSheet {
                sheet: *sheet,
                name: old_name,
                formulas: old_formulas,
            }];
            inverse.extend(defined_names);
            Ok(InvertedOp(inverse))
        }
        Op::SetDefinedNames { defined_names } => {
            let previous = std::mem::replace(&mut wb.defined_names, defined_names.clone());
            Ok(InvertedOp(vec![Op::SetDefinedNames {
                defined_names: previous,
            }]))
        }
        Op::InsertRows { sheet, at, count } => {
            apply_atomically(wb, |next| insert_rows(next, *sheet, *at, *count, op))
        }
        Op::DeleteRows { sheet, at, count } => {
            apply_atomically(wb, |next| delete_rows(next, *sheet, *at, *count, op))
        }
        Op::InsertCols { sheet, at, count } => {
            apply_atomically(wb, |next| insert_cols(next, *sheet, *at, *count, op))
        }
        Op::DeleteCols { sheet, at, count } => {
            apply_atomically(wb, |next| delete_cols(next, *sheet, *at, *count, op))
        }
    }
}

fn apply_atomically(
    wb: &mut Workbook,
    apply: impl FnOnce(&mut Workbook) -> Result<InvertedOp, OpError>,
) -> Result<InvertedOp, OpError> {
    let mut next = wb.clone();
    let inverse = apply(&mut next)?;
    *wb = next;
    Ok(inverse)
}

fn apply_range_formats(
    wb: &mut Workbook,
    sheet: SheetId,
    range: CellRange,
    mut update: impl FnMut(&mut xlsx_model::CellFormat, u32, u32) -> Result<(), OpError>,
) -> Result<InvertedOp, OpError> {
    let sheet_ref = wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let mut cells = Vec::new();
    for row in range.start.row..=range.end.row {
        for col in range.start.col..=range.end.col {
            let at = CellRef::new(row, col);
            cells.push((at, sheet_ref.cell(at).cloned().unwrap_or_default()));
        }
    }
    let mut inverse = Vec::with_capacity(cells.len());
    for (at, old) in cells {
        let mut format = wb.styles.cell_format(old.style);
        update(&mut format, at.row, at.col)?;
        let mut next = old.clone();
        next.style = wb.styles.intern_cell_format(&format);
        if next != old {
            wb.sheet_mut(sheet)
                .ok_or(OpError::SheetNotFound(sheet))?
                .set_cell(at, next);
            inverse.push(Op::SetCell {
                sheet,
                at,
                cell: CellState::from(old),
            });
        }
    }
    Ok(InvertedOp(inverse))
}

/// apply a sequence of ops, returning the combined inverse (per-op inverses
/// concatenated in reverse order).
pub fn apply_ops(wb: &mut Workbook, ops: &[Op]) -> Result<Vec<Op>, OpError> {
    let mut next = wb.clone();
    let inverse = apply_ops_in_place(&mut next, ops)?;
    *wb = next;
    Ok(inverse)
}

fn apply_ops_in_place(wb: &mut Workbook, ops: &[Op]) -> Result<Vec<Op>, OpError> {
    let mut per_op: Vec<Vec<Op>> = Vec::with_capacity(ops.len());
    for op in ops {
        per_op.push(apply(wb, op)?.0);
    }
    let mut inverse = Vec::new();
    for chunk in per_op.into_iter().rev() {
        inverse.extend(chunk);
    }
    Ok(inverse)
}

fn ranges_intersect(left: CellRange, right: CellRange) -> bool {
    left.start.row <= right.end.row
        && left.end.row >= right.start.row
        && left.start.col <= right.end.col
        && left.end.col >= right.start.col
}

fn range_contains(outer: CellRange, inner: CellRange) -> bool {
    outer.start.row <= inner.start.row
        && outer.end.row >= inner.end.row
        && outer.start.col <= inner.start.col
        && outer.end.col >= inner.end.col
}

/// a sort moves whole rows while merges stay fixed, so a merge is only safe
/// when it lies inside the range, spans a single row, and the same column
/// span is merged on every sorted row; anything else is a conflict.
fn find_sort_merge_conflict(s: &Sheet, range: CellRange) -> Option<CellRange> {
    let contained: Vec<CellRange> = s
        .merges
        .iter()
        .copied()
        .filter(|merge| ranges_intersect(*merge, range))
        .collect();
    contained
        .iter()
        .copied()
        .find(|merge| !range_contains(range, *merge) || merge.start.row != merge.end.row)
        .or_else(|| {
            contained.iter().copied().find(|merge| {
                (range.start.row..=range.end.row).any(|row| {
                    !contained.iter().any(|other| {
                        other.start.row == row
                            && other.start.col == merge.start.col
                            && other.end.col == merge.end.col
                    })
                })
            })
        })
}

/// the pieces of a row move's inverse: the reversed permutation plus `SetCell`
/// ops restoring the exact pre-move text of every rewritten formula.
#[derive(Default)]
struct RowMoveOutcome {
    inverse_moves: Vec<(RowId, RowId)>,
    formula_restores: Vec<Op>,
}

impl RowMoveOutcome {
    fn into_inverse(self, sheet: SheetId, range: CellRange) -> Vec<Op> {
        if self.inverse_moves.is_empty() {
            return Vec::new();
        }
        let mut inverse = vec![Op::MoveRows {
            sheet,
            range,
            moves: self.inverse_moves,
        }];
        inverse.extend(self.formula_restores);
        inverse
    }
}

/// move whole rows of cells within `range`'s columns; every `(from, to)` pair
/// is applied simultaneously, so sources are read before any destination is
/// written. cells carry their style and comment; formulas move with excel's
/// copy semantics — relative row references shift by the row's delta — and
/// hyperlinks follow their row. conflicts are detected before any mutation.
fn move_row_cells(
    s: &mut Sheet,
    sheet: SheetId,
    range: CellRange,
    moves: &[(RowId, RowId)],
) -> Result<RowMoveOutcome, OpError> {
    if moves.is_empty() {
        return Ok(RowMoveOutcome::default());
    }
    if let Some(hyperlink) = find_row_move_hyperlink_conflict(s, range, moves) {
        return Err(OpError::HyperlinkConflict { sheet, hyperlink });
    }
    let cols = range.start.col..range.end.col.saturating_add(1);
    let mut writes = Vec::with_capacity(moves.len());
    let mut formula_restores = Vec::new();
    for &(from, to) in moves {
        let delta = i64::from(to) - i64::from(from);
        let mut cells: Vec<(ColId, Cell)> = s
            .iter_cells_in_rect(from..from.saturating_add(1), cols.clone())
            .map(|(at, cell)| (at.col, cell.clone()))
            .collect();
        for (col, cell) in &mut cells {
            let Some(source) = &cell.formula else {
                continue;
            };
            let at = CellRef::new(from, *col);
            let rewritten = shift_formula_rows(source, delta)
                .map_err(|_| OpError::FormulaNotRewritable { sheet, cell: at })?;
            if let Some(rewritten) = rewritten {
                formula_restores.push(Op::SetCell {
                    sheet,
                    at,
                    cell: CellState::from(&*cell),
                });
                cell.formula = Some(rewritten);
            }
        }
        writes.push((to, cells));
    }
    let mut cleared = Vec::new();
    for &(_, to) in moves {
        cleared.extend(
            s.iter_cells_in_rect(to..to.saturating_add(1), cols.clone())
                .map(|(at, _)| at),
        );
    }
    for at in cleared {
        s.set_cell(at, Cell::default());
    }
    for (to, cells) in writes {
        for (col, cell) in cells {
            s.set_cell(CellRef::new(to, col), cell);
        }
    }
    move_row_comments(s, range, moves);
    move_row_hyperlinks(s, range, moves);
    Ok(RowMoveOutcome {
        inverse_moves: moves.iter().map(|&(from, to)| (to, from)).collect(),
        formula_restores,
    })
}

/// a hyperlink follows a row move only when it spans a single row, lies inside
/// the moved columns, and its row is a move source; any other hyperlink
/// touching a moved or overwritten row is a conflict.
fn find_row_move_hyperlink_conflict(
    s: &Sheet,
    range: CellRange,
    moves: &[(RowId, RowId)],
) -> Option<CellRange> {
    let mut affected = BTreeSet::new();
    for &(from, to) in moves {
        affected.insert(from);
        affected.insert(to);
    }
    s.hyperlinks.iter().map(|h| h.range).find(|r| {
        let rows_touch = affected.range(r.start.row..=r.end.row).next().is_some();
        let cols_touch = r.start.col <= range.end.col && r.end.col >= range.start.col;
        if !rows_touch || !cols_touch {
            return false;
        }
        let single_row = r.start.row == r.end.row;
        let inside_cols = r.start.col >= range.start.col && r.end.col <= range.end.col;
        !(single_row && inside_cols && moves.iter().any(|&(from, _)| from == r.start.row))
    })
}

/// permute single-row hyperlinks alongside their rows under the same
/// simultaneous semantics as `move_row_cells`; conflicting hyperlinks were
/// rejected before any mutation.
fn move_row_hyperlinks(s: &mut Sheet, range: CellRange, moves: &[(RowId, RowId)]) {
    let destinations: BTreeMap<RowId, RowId> = moves.iter().copied().collect();
    for hyperlink in &mut s.hyperlinks {
        let r = hyperlink.range;
        if r.start.row != r.end.row || r.start.col < range.start.col || r.end.col > range.end.col {
            continue;
        }
        if let Some(&to) = destinations.get(&r.start.row) {
            hyperlink.range.start.row = to;
            hyperlink.range.end.row = to;
        }
    }
}

/// permute comments alongside their rows under the same simultaneous
/// semantics as `move_row_cells`: read every source row's comments within
/// `range`'s columns, clear the destinations, then write.
fn move_row_comments(s: &mut Sheet, range: CellRange, moves: &[(RowId, RowId)]) {
    let mut writes = Vec::with_capacity(moves.len());
    for &(from, to) in moves {
        let comments: Vec<(ColId, Comment)> = s
            .comments
            .range((from, range.start.col)..=(from, range.end.col))
            .map(|(&(_, col), comment)| (col, comment.clone()))
            .collect();
        writes.push((to, comments));
    }
    for &(_, to) in moves {
        let cleared: Vec<(RowId, ColId)> = s
            .comments
            .range((to, range.start.col)..=(to, range.end.col))
            .map(|(&key, _)| key)
            .collect();
        for key in cleared {
            s.comments.remove(&key);
        }
    }
    for (to, comments) in writes {
        for (col, comment) in comments {
            s.comments.insert((to, col), comment);
        }
    }
}

/// sort-key ordering: empties sort last regardless of direction; otherwise
/// excel's collation approximation — numbers < text < booleans < errors.
fn compare_sort_keys(a: Option<&CellValue>, b: Option<&CellValue>, ascending: bool) -> Ordering {
    fn nonempty(value: Option<&CellValue>) -> Option<&CellValue> {
        value.filter(|value| !matches!(value, CellValue::Empty))
    }
    match (nonempty(a), nonempty(b)) {
        (None, None) => Ordering::Equal,
        (None, Some(_)) => Ordering::Greater,
        (Some(_), None) => Ordering::Less,
        (Some(a), Some(b)) => {
            let ordering = compare_values(a, b);
            if ascending {
                ordering
            } else {
                ordering.reverse()
            }
        }
    }
}

fn compare_values(a: &CellValue, b: &CellValue) -> Ordering {
    match (a, b) {
        (CellValue::Number { value: a }, CellValue::Number { value: b }) => a.total_cmp(b),
        (CellValue::Text { value: a }, CellValue::Text { value: b }) => a
            .to_lowercase()
            .cmp(&b.to_lowercase())
            .then_with(|| a.cmp(b)),
        (CellValue::Bool { value: a }, CellValue::Bool { value: b }) => a.cmp(b),
        (CellValue::Error { value: a }, CellValue::Error { value: b }) => {
            a.as_str().cmp(b.as_str())
        }
        _ => value_rank(a).cmp(&value_rank(b)),
    }
}

fn value_rank(value: &CellValue) -> u8 {
    match value {
        CellValue::Number { .. } => 0,
        CellValue::Text { .. } => 1,
        CellValue::Bool { .. } => 2,
        CellValue::Error { .. } => 3,
        CellValue::Empty => 4,
    }
}

/// the non-header rows of `filter`'s range that fail its criteria against the
/// sheet's current cells — the rows the filter itself hides.
fn filter_failing_rows(sheet: &Sheet, filter: Option<&AutoFilter>) -> BTreeSet<RowId> {
    let Some(filter) = filter else {
        return BTreeSet::new();
    };
    (filter.range.start.row..=filter.range.end.row)
        .filter(|&row| row != filter.range.start.row && !row_passes_filter(sheet, filter, row))
        .collect()
}

/// a row is visible when every column with value criteria matches its cell
/// text, blanks counting only when the column keeps blanks.
fn row_passes_filter(sheet: &Sheet, filter: &AutoFilter, row: RowId) -> bool {
    filter.columns.iter().all(|column| {
        let Some(values) = &column.values else {
            return true;
        };
        let text = sheet
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

/// the text a filter matches against: the bare value, no number formats.
fn filter_text(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number { value } => value.to_string(),
        CellValue::Text { value } => value.clone(),
        CellValue::Bool { value } => if *value { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::Error { value } => value.as_str().to_string(),
    }
}

/// remap an address through a structural op; `None` when it falls inside a
/// deleted span. non-structural ops leave the address unchanged.
pub fn remap_ref(at: CellRef, structural_op: &Op) -> Option<CellRef> {
    match *structural_op {
        Op::InsertRows { at: p, count, .. } => {
            shift_row(at, |r| insert_axis(r, p, count, MAX_ROWS))
        }
        Op::DeleteRows { at: p, count, .. } => shift_row(at, |r| delete_axis(r, p, count)),
        Op::InsertCols { at: p, count, .. } => {
            shift_col(at, |c| insert_axis(c, p, count, MAX_COLS))
        }
        Op::DeleteCols { at: p, count, .. } => shift_col(at, |c| delete_axis(c, p, count)),
        _ => Some(at),
    }
}

/// remap both corners of a range; `None` if either corner is deleted.
fn remap_range(range: CellRange, op: &Op) -> Option<CellRange> {
    let start = remap_ref(range.start, op)?;
    let end = remap_ref(range.end, op)?;
    Some(CellRange::new(start, end))
}

/// index shift under an insert at `p` of `count`. indices at or after `p` move
/// up by `count`; anything pushed past `limit` is dropped (`None`).
fn insert_axis(idx: u32, p: u32, count: u32, limit: u32) -> Option<u32> {
    if idx < p {
        return Some(idx);
    }
    let shifted = u64::from(idx) + u64::from(count);
    if shifted > u64::from(limit - 1) {
        None
    } else {
        Some(shifted as u32)
    }
}

/// index shift under a delete at `p` of `count`. indices inside `[p, p+count)`
/// are deleted (`None`); later indices move down by `count`.
fn delete_axis(idx: u32, p: u32, count: u32) -> Option<u32> {
    if idx < p {
        Some(idx)
    } else if idx < p.saturating_add(count) {
        None
    } else {
        Some(idx - count)
    }
}

/// shift an inclusive index span through an insert at `p`; ends pushed past
/// the axis limit are clamped to it.
fn insert_span(start: u32, end: u32, p: u32, count: u32, limit: u32) -> (u32, u32) {
    let shift = |idx| insert_axis(idx, p, count, limit).unwrap_or(limit - 1);
    (shift(start), shift(end))
}

/// shift an inclusive index span through a delete at `p` of `count`; `None`
/// when the whole span is deleted, otherwise the surviving span.
fn delete_span(start: u32, end: u32, p: u32, count: u32) -> Option<(u32, u32)> {
    let cutoff = p.saturating_add(count);
    let start = if start >= cutoff {
        start - count
    } else if start >= p {
        p
    } else {
        start
    };
    let end = if end >= cutoff {
        end - count
    } else if end >= p {
        p.checked_sub(1)?
    } else {
        end
    };
    (start <= end).then_some((start, end))
}

fn shift_row(mut at: CellRef, f: impl Fn(u32) -> Option<u32>) -> Option<CellRef> {
    at.row = f(at.row)?;
    Some(at)
}

fn shift_col(mut at: CellRef, f: impl Fn(u32) -> Option<u32>) -> Option<CellRef> {
    at.col = f(at.col)?;
    Some(at)
}

fn insert_rows(
    wb: &mut Workbook,
    sheet: SheetId,
    at: RowId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let restores = remap_formulas(wb, op)?;
    let defined_name_restore = remap_defined_names(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let old_filter = s.auto_filter.clone();
    let old_hidden = s.hidden_rows.clone();
    let dropped = shift_cells(s, op);
    shift_row_heights_up(s, at, count);
    remap_merges_keep(s, op);
    remap_hyperlinks(s, op);
    remap_auto_filter(s, op);
    remap_hidden_rows(s, op);
    let dropped_comments = remap_comments(s, op);

    let mut inv = vec![Op::DeleteRows { sheet, at, count }];
    for (r, c) in dropped {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    inv.extend(comment_restore_ops(sheet, dropped_comments));
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(filter_restore_ops(s, sheet, old_filter, old_hidden));
    inv.extend(restores);
    inv.extend(defined_name_restore);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn delete_rows(
    wb: &mut Workbook,
    sheet: SheetId,
    at: RowId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let restores = remap_formulas(wb, op)?;
    let defined_name_restore = remap_defined_names(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let old_filter = s.auto_filter.clone();
    let old_hidden = s.hidden_rows.clone();
    let deleted = shift_cells(s, op);
    let dropped_heights = shift_row_heights_down(s, at, count);
    let dropped_merges = remap_merges_drop(s, op);
    remap_hyperlinks(s, op);
    remap_auto_filter(s, op);
    remap_hidden_rows(s, op);
    let dropped_comments = remap_comments(s, op);

    let mut inv = vec![Op::InsertRows { sheet, at, count }];
    for (r, c) in deleted {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    for (row, h) in dropped_heights {
        inv.push(Op::SetRowHeight {
            sheet,
            row,
            height: Some(h),
        });
    }
    for range in dropped_merges {
        inv.push(Op::MergeCells { sheet, range });
    }
    inv.extend(comment_restore_ops(sheet, dropped_comments));
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(filter_restore_ops(s, sheet, old_filter, old_hidden));
    inv.extend(restores);
    inv.extend(defined_name_restore);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn insert_cols(
    wb: &mut Workbook,
    sheet: SheetId,
    at: ColId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let restores = remap_formulas(wb, op)?;
    let defined_name_restore = remap_defined_names(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let old_filter = s.auto_filter.clone();
    let old_hidden = s.hidden_rows.clone();
    let dropped = shift_cells(s, op);
    shift_col_widths_up(s, at, count);
    remap_merges_keep(s, op);
    remap_hyperlinks(s, op);
    remap_auto_filter(s, op);
    let dropped_comments = remap_comments(s, op);

    let mut inv = vec![Op::DeleteCols { sheet, at, count }];
    for (r, c) in dropped {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    inv.extend(comment_restore_ops(sheet, dropped_comments));
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(filter_restore_ops(s, sheet, old_filter, old_hidden));
    inv.extend(restores);
    inv.extend(defined_name_restore);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

fn delete_cols(
    wb: &mut Workbook,
    sheet: SheetId,
    at: ColId,
    count: u32,
    op: &Op,
) -> Result<InvertedOp, OpError> {
    wb.sheet(sheet).ok_or(OpError::SheetNotFound(sheet))?;
    let restores = remap_formulas(wb, op)?;
    let defined_name_restore = remap_defined_names(wb, op)?;
    let hyperlink_restores = remap_hyperlink_locations(wb, op);
    let s = sheet_mut(wb, sheet)?;
    let old_hyperlinks = s.hyperlinks.clone();
    let old_filter = s.auto_filter.clone();
    let old_hidden = s.hidden_rows.clone();
    let deleted = shift_cells(s, op);
    let dropped_widths = shift_col_widths_down(s, at, count);
    let dropped_merges = remap_merges_drop(s, op);
    remap_hyperlinks(s, op);
    remap_auto_filter(s, op);
    let dropped_comments = remap_comments(s, op);

    let mut inv = vec![Op::InsertCols { sheet, at, count }];
    for (r, c) in deleted {
        inv.push(Op::SetCell {
            sheet,
            at: r,
            cell: CellState::from(&c),
        });
    }
    for (col, w) in dropped_widths {
        inv.push(Op::SetColWidth {
            sheet,
            col,
            width: Some(w),
        });
    }
    for range in dropped_merges {
        inv.push(Op::MergeCells { sheet, range });
    }
    inv.extend(comment_restore_ops(sheet, dropped_comments));
    inv.push(Op::SetHyperlinks {
        sheet,
        hyperlinks: old_hyperlinks,
    });
    inv.extend(filter_restore_ops(s, sheet, old_filter, old_hidden));
    inv.extend(restores);
    inv.extend(defined_name_restore);
    inv.extend(hyperlink_restores);
    Ok(InvertedOp(inv))
}

/// remap every occupied cell through `op`, rebuilding storage. returns the
/// cells whose address was dropped, for the inverse.
fn shift_cells(s: &mut Sheet, op: &Op) -> Vec<(CellRef, Cell)> {
    let old: Vec<(CellRef, Cell)> = s.iter_cells().map(|(r, c)| (r, c.clone())).collect();
    let mut moved = Vec::new();
    let mut dropped = Vec::new();
    for (r, c) in &old {
        match remap_ref(*r, op) {
            Some(nr) => moved.push((nr, c.clone())),
            None => dropped.push((*r, c.clone())),
        }
    }
    for (r, _) in &old {
        s.set_cell(*r, Cell::default());
    }
    for (nr, c) in moved {
        s.set_cell(nr, c);
    }
    dropped
}

fn shift_row_heights_up(s: &mut Sheet, at: RowId, count: u32) {
    let shifted: BTreeMap<RowId, f64> = s
        .row_heights
        .iter()
        .map(|(&row, &h)| {
            (
                if row >= at {
                    row.saturating_add(count)
                } else {
                    row
                },
                h,
            )
        })
        .collect();
    s.row_heights = shifted;
}

fn shift_row_heights_down(s: &mut Sheet, at: RowId, count: u32) -> Vec<(RowId, f64)> {
    let mut kept = BTreeMap::new();
    let mut dropped = Vec::new();
    for (&row, &h) in &s.row_heights {
        if row < at {
            kept.insert(row, h);
        } else if row >= at.saturating_add(count) {
            kept.insert(row - count, h);
        } else {
            dropped.push((row, h));
        }
    }
    s.row_heights = kept;
    dropped
}

fn shift_col_widths_up(s: &mut Sheet, at: ColId, count: u32) {
    let shifted: BTreeMap<ColId, f64> = s
        .col_widths
        .iter()
        .map(|(&col, &w)| {
            (
                if col >= at {
                    col.saturating_add(count)
                } else {
                    col
                },
                w,
            )
        })
        .collect();
    s.col_widths = shifted;
}

fn shift_col_widths_down(s: &mut Sheet, at: ColId, count: u32) -> Vec<(ColId, f64)> {
    let mut kept = BTreeMap::new();
    let mut dropped = Vec::new();
    for (&col, &w) in &s.col_widths {
        if col < at {
            kept.insert(col, w);
        } else if col >= at.saturating_add(count) {
            kept.insert(col - count, w);
        } else {
            dropped.push((col, w));
        }
    }
    s.col_widths = kept;
    dropped
}

/// remap merges under an insert (no corner is ever deleted).
fn remap_merges_keep(s: &mut Sheet, op: &Op) {
    let remapped: Vec<CellRange> = s
        .merges
        .iter()
        .filter_map(|m| remap_range(*m, op))
        .collect();
    s.merges = remapped;
}

/// remap merges under a delete; any merge with a deleted corner is dropped and
/// returned for the inverse to restore.
fn remap_merges_drop(s: &mut Sheet, op: &Op) -> Vec<CellRange> {
    let mut kept = Vec::new();
    let mut dropped = Vec::new();
    for m in &s.merges {
        match remap_range(*m, op) {
            Some(nm) => kept.push(nm),
            None => dropped.push(*m),
        }
    }
    s.merges = kept;
    dropped
}

/// remap explicitly hidden rows through a row insert or delete, dropping
/// entries for deleted rows.
fn remap_hidden_rows(sheet: &mut Sheet, op: &Op) {
    sheet.hidden_rows = sheet
        .hidden_rows
        .iter()
        .filter_map(|&row| match *op {
            Op::InsertRows { at, count, .. } => insert_axis(row, at, count, MAX_ROWS),
            Op::DeleteRows { at, count, .. } => delete_axis(row, at, count),
            _ => Some(row),
        })
        .collect();
}

/// remap comment anchors through a structural op, returning the comments
/// whose cell was deleted so the inverse can restore them.
fn remap_comments(sheet: &mut Sheet, op: &Op) -> Vec<(CellRef, Comment)> {
    let old = std::mem::take(&mut sheet.comments);
    let mut dropped = Vec::new();
    for ((row, col), comment) in old {
        let at = CellRef::new(row, col);
        match remap_ref(at, op) {
            Some(moved) => {
                sheet.comments.insert((moved.row, moved.col), comment);
            }
            None => dropped.push((at, comment)),
        }
    }
    dropped
}

fn comment_restore_ops(sheet: SheetId, dropped: Vec<(CellRef, Comment)>) -> Vec<Op> {
    dropped
        .into_iter()
        .map(|(cell, comment)| Op::SetComment {
            sheet,
            cell,
            comment: Some(comment),
        })
        .collect()
}

/// remap the auto filter through a structural op: its range shifts with the
/// edited axis, deleted criteria columns are dropped, and the filter itself
/// is dropped when its whole range is deleted.
fn remap_auto_filter(sheet: &mut Sheet, op: &Op) {
    let Some(mut filter) = sheet.auto_filter.take() else {
        return;
    };
    match *op {
        Op::InsertRows { at, count, .. } => {
            let (start, end) = insert_span(
                filter.range.start.row,
                filter.range.end.row,
                at,
                count,
                MAX_ROWS,
            );
            filter.range.start.row = start;
            filter.range.end.row = end;
        }
        Op::DeleteRows { at, count, .. } => {
            let Some((start, end)) =
                delete_span(filter.range.start.row, filter.range.end.row, at, count)
            else {
                return;
            };
            filter.range.start.row = start;
            filter.range.end.row = end;
        }
        Op::InsertCols { at, count, .. } => {
            let (start, end) = insert_span(
                filter.range.start.col,
                filter.range.end.col,
                at,
                count,
                MAX_COLS,
            );
            filter.range.start.col = start;
            filter.range.end.col = end;
            filter.columns.retain_mut(|column| {
                insert_axis(column.col, at, count, MAX_COLS).is_some_and(|col| {
                    column.col = col;
                    true
                })
            });
        }
        Op::DeleteCols { at, count, .. } => {
            let Some((start, end)) =
                delete_span(filter.range.start.col, filter.range.end.col, at, count)
            else {
                return;
            };
            filter.range.start.col = start;
            filter.range.end.col = end;
            filter.columns.retain_mut(|column| {
                delete_axis(column.col, at, count).is_some_and(|col| {
                    column.col = col;
                    true
                })
            });
        }
        _ => {}
    }
    sheet.auto_filter = Some(filter);
}

/// inverse ops restoring the exact pre-edit filter and hidden set; the
/// `SetHiddenRows` follows `SetAutoFilter` so its replayed row evaluation is
/// overridden by the recorded state.
fn filter_restore_ops(
    sheet: &Sheet,
    id: SheetId,
    old_filter: Option<AutoFilter>,
    old_hidden: BTreeSet<RowId>,
) -> Vec<Op> {
    if sheet.auto_filter == old_filter && sheet.hidden_rows == old_hidden {
        return Vec::new();
    }
    vec![
        Op::SetAutoFilter {
            sheet: id,
            filter: old_filter,
        },
        Op::SetHiddenRows {
            sheet: id,
            hidden: old_hidden.into_iter().collect(),
        },
    ]
}

fn remap_hyperlinks(sheet: &mut Sheet, op: &Op) {
    sheet.hyperlinks = sheet
        .hyperlinks
        .drain(..)
        .filter_map(|mut hyperlink| {
            hyperlink.range = remap_hyperlink_range(hyperlink.range, op)?;
            Some(hyperlink)
        })
        .collect();
}

fn rename_hyperlink_locations(wb: &mut Workbook, old_name: &str, new_name: &str) -> Vec<Op> {
    let mut restores = Vec::new();
    for (index, sheet) in wb.sheets.iter_mut().enumerate() {
        let mut changed = false;
        let mut hyperlinks = sheet.hyperlinks.clone();
        for hyperlink in &mut hyperlinks {
            let Some(location) = &hyperlink.location else {
                continue;
            };
            let rewritten = rename_hyperlink_location(location, old_name, new_name);
            if rewritten != *location {
                hyperlink.location = Some(rewritten);
                changed = true;
            }
        }
        if changed {
            let old = std::mem::replace(&mut sheet.hyperlinks, hyperlinks);
            restores.push(Op::SetHyperlinks {
                sheet: SheetId(index as u32),
                hyperlinks: old,
            });
        }
    }
    restores
}

fn remove_sheet(wb: &mut Workbook, index: usize) -> Result<InvertedOp, OpError> {
    if index >= wb.sheets.len() {
        return Err(OpError::SheetIndexOutOfRange(index));
    }
    let removed = wb.sheets.remove(index);
    let previous_defined_names = drop_defined_name_scopes(wb, index);
    let sheet = SheetId(index as u32);
    let mut inv = vec![Op::AddSheet {
        index,
        name: removed.name.clone(),
    }];
    for (at, cell) in removed.iter_cells() {
        inv.push(Op::SetCell {
            sheet,
            at,
            cell: CellState::from(cell),
        });
    }
    for range in &removed.merges {
        inv.push(Op::MergeCells {
            sheet,
            range: *range,
        });
    }
    for (&col, &w) in &removed.col_widths {
        inv.push(Op::SetColWidth {
            sheet,
            col,
            width: Some(w),
        });
    }
    for (&row, &h) in &removed.row_heights {
        inv.push(Op::SetRowHeight {
            sheet,
            row,
            height: Some(h),
        });
    }
    if let Some(pane) = removed.freeze_pane {
        inv.push(Op::SetFreezePane {
            sheet,
            pane: Some(pane),
        });
    }
    if let Some(filter) = &removed.auto_filter {
        inv.push(Op::SetAutoFilter {
            sheet,
            filter: Some(filter.clone()),
        });
    }
    if !removed.hidden_rows.is_empty() {
        inv.push(Op::SetHiddenRows {
            sheet,
            hidden: removed.hidden_rows.iter().copied().collect(),
        });
    }
    if !removed.hyperlinks.is_empty() {
        inv.push(Op::SetHyperlinks {
            sheet,
            hyperlinks: removed.hyperlinks,
        });
    }
    for ((row, col), comment) in removed.comments {
        inv.push(Op::SetComment {
            sheet,
            cell: CellRef::new(row, col),
            comment: Some(comment),
        });
    }
    if let Some(defined_names) = previous_defined_names {
        inv.push(Op::SetDefinedNames { defined_names });
    }
    Ok(InvertedOp(inv))
}

/// Widens sheet scopes at or after an inserted sheet.
fn shift_defined_name_scopes(wb: &mut Workbook, index: usize) {
    for defined in &mut wb.defined_names {
        if let Some(sheet) = defined.local_sheet
            && sheet.0 as usize >= index
        {
            defined.local_sheet = Some(SheetId(sheet.0 + 1));
        }
    }
}

/// Drops names scoped to a removed sheet and narrows the scopes above it,
/// returning the prior list when it changed.
fn drop_defined_name_scopes(wb: &mut Workbook, index: usize) -> Option<Vec<DefinedName>> {
    let previous = wb.defined_names.clone();
    wb.defined_names
        .retain(|defined| defined.local_sheet.map(|sheet| sheet.0 as usize) != Some(index));
    for defined in &mut wb.defined_names {
        if let Some(sheet) = defined.local_sheet
            && sheet.0 as usize > index
        {
            defined.local_sheet = Some(SheetId(sheet.0 - 1));
        }
    }
    (previous != wb.defined_names).then_some(previous)
}

fn sheet_mut(wb: &mut Workbook, sheet: SheetId) -> Result<&mut Sheet, OpError> {
    wb.sheet_mut(sheet).ok_or(OpError::SheetNotFound(sheet))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{AutoFilterColumn, CellProvider, CellValue, FreezePane, Hyperlink};

    fn r(a1: &str) -> CellRef {
        CellRef::parse_a1(a1).unwrap()
    }

    fn rng(a1: &str) -> CellRange {
        CellRange::parse_a1(a1).unwrap()
    }

    fn text_cell(value: &str) -> Cell {
        Cell {
            value: CellValue::Text {
                value: value.into(),
            },
            ..Cell::default()
        }
    }

    fn num(v: f64) -> CellState {
        CellState {
            value: CellValue::Number { value: v },
            ..Default::default()
        }
    }

    fn wb_one_sheet() -> Workbook {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        wb
    }

    #[test]
    fn remap_ref_insert_rows_edges() {
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 5,
            count: 3,
        };
        // a1 row N is 0-based index N-1, so A5 sits before the insert at 5
        assert_eq!(remap_ref(r("A5"), &op), Some(r("A5")));
        assert_eq!(remap_ref(r("A6"), &op).unwrap().row, 5 + 3);
        assert_eq!(remap_ref(r("A100"), &op).unwrap().row, 99 + 3);
    }

    #[test]
    fn remap_ref_delete_rows_edges() {
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 5,
            count: 3,
        };
        assert_eq!(remap_ref(r("A5"), &op), Some(r("A5")));
        assert_eq!(remap_ref(CellRef::new(5, 0), &op), None);
        assert_eq!(remap_ref(CellRef::new(7, 0), &op), None);
        assert_eq!(remap_ref(CellRef::new(8, 0), &op), Some(CellRef::new(5, 0)));
    }

    #[test]
    fn remap_ref_cols_and_anchors() {
        let ins = Op::InsertCols {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        };
        let mut anchored = r("C1");
        anchored.abs_col = true;
        let out = remap_ref(anchored, &ins).unwrap();
        assert_eq!(out.col, 3);
        assert!(out.abs_col, "anchor preserved through remap");

        let del = Op::DeleteCols {
            sheet: SheetId(0),
            at: 2,
            count: 2,
        };
        assert_eq!(remap_ref(CellRef::new(0, 2), &del), None);
        assert_eq!(
            remap_ref(CellRef::new(0, 4), &del),
            Some(CellRef::new(0, 2))
        );
    }

    #[test]
    fn remap_ref_ignores_non_structural() {
        let op = Op::SetCell {
            sheet: SheetId(0),
            at: r("A1"),
            cell: num(1.0),
        };
        assert_eq!(remap_ref(r("Z9"), &op), Some(r("Z9")));
    }

    #[test]
    fn set_cell_round_trip() {
        let mut wb = wb_one_sheet();
        let op = Op::SetCell {
            sheet: SheetId(0),
            at: r("B2"),
            cell: num(42.0),
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("B2")),
            CellValue::Number { value: 42.0 }
        );
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.value(SheetId(0), r("B2")), CellValue::Empty);
        assert!(wb.sheet(SheetId(0)).unwrap().used_range().is_none());
    }

    #[test]
    fn delete_rows_restores_contents_and_merges() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(
            r("A1"),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Default::default()
            },
        );
        s.set_cell(
            r("A6"),
            Cell {
                value: CellValue::Number { value: 6.0 },
                ..Default::default()
            },
        );
        s.merges.push(CellRange::parse_a1("A6:B6").unwrap());
        s.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A6:B6").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: None,
        });
        s.row_heights.insert(5, 30.0);
        let before = wb.clone();

        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 2,
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A4")),
            CellValue::Number { value: 6.0 }
        );

        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        let after = &wb.sheets[0];
        let orig = &before.sheets[0];
        assert_eq!(after.used_range(), orig.used_range());
        assert_eq!(
            wb.value(SheetId(0), r("A6")),
            CellValue::Number { value: 6.0 }
        );
        assert_eq!(after.merges, orig.merges);
        assert_eq!(after.hyperlinks, orig.hyperlinks);
        assert_eq!(after.row_heights, orig.row_heights);
    }

    #[test]
    fn insert_rows_round_trip() {
        let mut wb = wb_one_sheet();
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            r("A3"),
            Cell {
                value: CellValue::Text { value: "x".into() },
                ..Default::default()
            },
        );
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .hyperlinks
            .push(Hyperlink {
                range: CellRange::parse_a1("A3:B3").unwrap(),
                external_target: None,
                location: Some("Sheet1!A1".into()),
                tooltip: None,
                display: None,
            });
        let before = wb.sheets[0].used_range();

        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 1,
            count: 2,
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(
            wb.value(SheetId(0), r("A5")),
            CellValue::Text { value: "x".into() }
        );
        assert_eq!(wb.sheets[0].hyperlinks[0].range.to_a1(), "A5:B5");

        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].used_range(), before);
        assert_eq!(
            wb.value(SheetId(0), r("A3")),
            CellValue::Text { value: "x".into() }
        );
        assert_eq!(wb.sheets[0].hyperlinks[0].range.to_a1(), "A3:B3");
    }

    #[test]
    fn structural_edits_and_renames_rewrite_hyperlink_destinations() {
        let mut wb = wb_one_sheet();
        wb.sheets[0].hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1").unwrap(),
            external_target: Some("https://example.com/report".into()),
            location: Some("Target!A3".into()),
            tooltip: None,
            display: Some("Jump".into()),
        });
        wb.sheets.push(Sheet::new("Target"));
        wb.sheets[1].hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1:A4").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: None,
        });
        let before = wb.clone();

        let inverse = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(1),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("Target!A5")
        );
        assert_eq!(
            wb.sheets[0].hyperlinks[0].external_target.as_deref(),
            Some("https://example.com/report")
        );
        assert_eq!(wb.sheets[1].hyperlinks[0].range.to_a1(), "A1:A6");
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);

        let inverse = apply(
            &mut wb,
            &Op::DeleteRows {
                sheet: SheetId(1),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[1].hyperlinks[0].range.to_a1(), "A1:A3");
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("Target!A2")
        );
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);

        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(1),
                name: "New Target".into(),
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("'New Target'!A3")
        );
        assert_eq!(
            wb.sheets[0].hyperlinks[0].external_target.as_deref(),
            Some("https://example.com/report")
        );
        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);
    }

    #[test]
    fn rename_rewrites_hash_prefixed_hyperlink_locations() {
        let mut wb = wb_one_sheet();
        let link = |range: &str, location: &str| Hyperlink {
            range: CellRange::parse_a1(range).unwrap(),
            external_target: None,
            location: Some(location.into()),
            tooltip: None,
            display: None,
        };
        wb.sheets[0].hyperlinks.extend([
            link("A1", "#Target!A1"),
            link("A2", "Target!A2"),
            link("A3", "#'Target'!A3"),
            link("A4", "#MyRange"),
            link("A5", "#Other!A1"),
        ]);
        wb.sheets[0].hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A6").unwrap(),
            external_target: Some("https://example.com/report".into()),
            location: Some("#Target!A6".into()),
            tooltip: None,
            display: None,
        });
        wb.sheets.push(Sheet::new("Target"));
        let before = wb.clone();

        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(1),
                name: "My Sheet".into(),
            },
        )
        .unwrap();

        let locations: Vec<Option<&str>> = wb.sheets[0]
            .hyperlinks
            .iter()
            .map(|hyperlink| hyperlink.location.as_deref())
            .collect();
        assert_eq!(
            locations,
            vec![
                Some("#'My Sheet'!A1"),
                Some("'My Sheet'!A2"),
                Some("#'My Sheet'!A3"),
                Some("#MyRange"),
                Some("#Other!A1"),
                Some("#'My Sheet'!A6"),
            ]
        );
        assert_eq!(
            wb.sheets[0].hyperlinks[5].external_target.as_deref(),
            Some("https://example.com/report")
        );

        for operation in &inverse.0 {
            apply(&mut wb, operation).unwrap();
        }
        assert_eq!(wb, before);
    }

    #[test]
    fn add_and_remove_sheet_round_trip() {
        let mut wb = wb_one_sheet();
        let op = Op::AddSheet {
            index: 1,
            name: "New".into(),
        };
        let inv = apply(&mut wb, &op).unwrap();
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[1].name, "New");
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets.len(), 1);
    }

    #[test]
    fn remove_populated_sheet_round_trips_contents() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Two"));
        let s = wb.sheet_mut(SheetId(1)).unwrap();
        s.set_cell(
            r("A1"),
            Cell {
                value: CellValue::Number { value: 9.0 },
                ..Default::default()
            },
        );
        s.merges.push(CellRange::parse_a1("A1:B1").unwrap());
        s.col_widths.insert(0, 12.0);
        s.freeze_pane = Some(FreezePane::new(1, 1, r("C4")));
        let before = wb.sheets[1].clone();

        let inv = apply(&mut wb, &Op::RemoveSheet { index: 1 }).unwrap();
        assert_eq!(wb.sheets.len(), 1);
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(
            wb.value(SheetId(1), r("A1")),
            CellValue::Number { value: 9.0 }
        );
        assert_eq!(wb.sheets[1].merges, before.merges);
        assert_eq!(wb.sheets[1].col_widths, before.col_widths);
        assert_eq!(wb.sheets[1].freeze_pane, before.freeze_pane);
    }

    #[test]
    fn merge_and_col_width_round_trip() {
        let mut wb = wb_one_sheet();
        let range = CellRange::parse_a1("A1:B2").unwrap();
        let inv = apply(
            &mut wb,
            &Op::MergeCells {
                sheet: SheetId(0),
                range,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].merges, vec![range]);
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert!(wb.sheets[0].merges.is_empty());

        let inv = apply(
            &mut wb,
            &Op::SetColWidth {
                sheet: SheetId(0),
                col: 3,
                width: Some(20.0),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].col_widths.get(&3), Some(&20.0));
        for iop in &inv.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].col_widths.get(&3), None);
    }

    #[test]
    fn set_freeze_pane_round_trip() {
        let mut wb = wb_one_sheet();
        let first = FreezePane::new(2, 1, r("B3"));
        let inv_set = apply(
            &mut wb,
            &Op::SetFreezePane {
                sheet: SheetId(0),
                pane: Some(first),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].freeze_pane, Some(first));

        let replacement = FreezePane::new(1, 0, r("A2"));
        let inv_replace = apply(
            &mut wb,
            &Op::SetFreezePane {
                sheet: SheetId(0),
                pane: Some(replacement),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].freeze_pane, Some(replacement));

        for iop in &inv_replace.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].freeze_pane, Some(first));
        for iop in &inv_set.0 {
            apply(&mut wb, iop).unwrap();
        }
        assert_eq!(wb.sheets[0].freeze_pane, None);
    }

    #[test]
    fn set_freeze_pane_missing_sheet_errors() {
        let mut wb = wb_one_sheet();
        let err = apply(
            &mut wb,
            &Op::SetFreezePane {
                sheet: SheetId(7),
                pane: None,
            },
        );
        assert_eq!(err.unwrap_err(), OpError::SheetNotFound(SheetId(7)));
    }

    #[test]
    fn rename_sheet_round_trip() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Future!A1+Sheet1!A1".into()),
                ..Cell::default()
            },
        );
        let inv = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].name, "Future");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Future!A1"));
        let redo = apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb.sheets[0].name, "Sheet1");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Sheet1!A1"));
        apply_ops(&mut wb, &redo).unwrap();
        assert_eq!(wb.sheets[0].name, "Future");
        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1+Future!A1"));
    }

    #[test]
    fn rename_undo_does_not_change_destination_only_references() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Future!A1".into()),
                ..Cell::default()
            },
        );
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A2"),
            Cell {
                formula: Some("Sheet1!A1".into()),
                ..Cell::default()
            },
        );

        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        match &inverse.0[0] {
            Op::RestoreSheet { formulas, .. } => assert_eq!(formulas.len(), 1),
            other => panic!("expected restore sheet inverse, got {other:?}"),
        }
        apply_ops(&mut wb, &inverse.0).unwrap();

        assert_eq!(wb.formula(SheetId(1), r("A1")), Some("Future!A1"));
        assert_eq!(wb.formula(SheetId(1), r("A2")), Some("Sheet1!A1"));
    }

    #[test]
    fn rename_history_stays_constant_across_undo_redo() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Formula"));
        wb.sheet_mut(SheetId(1)).unwrap().set_cell(
            r("A1"),
            Cell {
                formula: Some("Sheet1!A1".into()),
                ..Cell::default()
            },
        );
        let inverse = apply(
            &mut wb,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Future".into(),
            },
        )
        .unwrap();
        assert_eq!(inverse.0.len(), 1);
        let json = serde_json::to_string(&inverse.0).unwrap();
        let decoded: Vec<Op> = serde_json::from_str(&json).unwrap();
        assert_eq!(decoded, inverse.0);
        let redo = apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(redo.len(), 1);
        let undo = apply_ops(&mut wb, &redo).unwrap();
        assert_eq!(undo.len(), 1);
    }

    #[test]
    fn missing_sheet_errors() {
        let mut wb = Workbook::default();
        let err = apply(
            &mut wb,
            &Op::SetCell {
                sheet: SheetId(3),
                at: r("A1"),
                cell: num(1.0),
            },
        );
        assert_eq!(err.unwrap_err(), OpError::SheetNotFound(SheetId(3)));
    }

    #[test]
    fn apply_ops_is_atomic_when_formula_rewriting_fails() {
        let mut wb = wb_one_sheet();
        let before = wb.clone();
        let error = apply_ops(
            &mut wb,
            &[
                Op::SetCell {
                    sheet: SheetId(0),
                    at: r("A1"),
                    cell: CellState {
                        formula: Some("SUM(".into()),
                        ..CellState::default()
                    },
                },
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
            ],
        )
        .unwrap_err();
        assert!(matches!(error, OpError::FormulaNotRewritable { .. }));
        assert_eq!(wb, before);
    }

    #[test]
    fn structural_edits_capture_defined_name_rewrites_for_undo() {
        let mut wb = wb_one_sheet();
        wb.defined_names.push(DefinedName {
            name: "Input".into(),
            formula: "Sheet1!$A$2".into(),
            local_sheet: None,
            hidden: false,
        });
        let before = wb.clone();

        let inverse = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();

        assert_eq!(wb.defined_names[0].formula, "Sheet1!$A$3");
        assert!(
            inverse
                .0
                .iter()
                .any(|op| matches!(op, Op::SetDefinedNames { .. }))
        );
        apply_ops(&mut wb, &inverse.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn structural_edits_refuse_workbook_names_with_unqualified_references() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Other"));
        wb.defined_names.push(DefinedName {
            name: "Input".into(),
            formula: "$A$2".into(),
            local_sheet: None,
            hidden: false,
        });
        let before = wb.clone();

        let error = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap_err();

        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));
        assert_eq!(wb, before);
    }

    #[test]
    fn sort_range_orders_rows_and_undo_restores_original_order() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        for (row, value) in [(0, 3.0), (1, 1.0), (2, 2.0)] {
            s.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number { value },
                    ..Cell::default()
                },
            );
        }
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::SortRange {
                sheet: SheetId(0),
                range: rng("A1:A3"),
                key_col: 0,
                ascending: true,
            },
        )
        .unwrap();
        let values: Vec<CellValue> = (0..3)
            .map(|row| wb.value(SheetId(0), CellRef::new(row, 0)))
            .collect();
        assert_eq!(
            values,
            vec![
                CellValue::Number { value: 1.0 },
                CellValue::Number { value: 2.0 },
                CellValue::Number { value: 3.0 },
            ]
        );

        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn sort_range_moves_comments_with_their_rows_and_undo_restores_them() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        for (row, value) in [(0, 3.0), (1, 1.0), (2, 2.0)] {
            s.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number { value },
                    ..Cell::default()
                },
            );
        }
        s.set_comment(r("A1"), Some(note("Ada", "on the 3.0 row")));
        s.set_comment(r("B2"), Some(note("Grace", "on the 1.0 row")));
        s.set_comment(r("C1"), Some(note("Lin", "outside the sorted columns")));
        s.set_comment(r("A4"), Some(note("Mei", "below the sorted rows")));
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::SortRange {
                sheet: SheetId(0),
                range: rng("A1:B3"),
                key_col: 0,
                ascending: true,
            },
        )
        .unwrap();
        let sheet = &wb.sheets[0];
        assert_eq!(
            sheet.comment_at(r("A3")),
            Some(&note("Ada", "on the 3.0 row"))
        );
        assert_eq!(
            sheet.comment_at(r("B1")),
            Some(&note("Grace", "on the 1.0 row"))
        );
        assert_eq!(sheet.comment_at(r("A1")), None);
        assert_eq!(sheet.comment_at(r("B2")), None);
        assert_eq!(
            sheet.comment_at(r("C1")),
            Some(&note("Lin", "outside the sorted columns"))
        );
        assert_eq!(
            sheet.comment_at(r("A4")),
            Some(&note("Mei", "below the sorted rows"))
        );
        assert_eq!(sheet.comments.len(), 4);

        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn move_rows_swaps_comments_between_commented_and_uncommented_rows() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(r("A1"), text_cell("first"));
        s.set_cell(r("A2"), text_cell("second"));
        s.set_comment(r("A2"), Some(note("Ada", "swap me")));
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::MoveRows {
                sheet: SheetId(0),
                range: rng("A1:A2"),
                moves: vec![(0, 1), (1, 0)],
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].comment_at(r("A1")),
            Some(&note("Ada", "swap me"))
        );
        assert_eq!(wb.sheets[0].comment_at(r("A2")), None);

        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn set_auto_filter_hides_non_matching_rows_but_never_the_header() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(r("A1"), text_cell("Name"));
        s.set_cell(r("A2"), text_cell("keep"));
        s.set_cell(r("A3"), text_cell("drop"));
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::SetAutoFilter {
                sheet: SheetId(0),
                filter: Some(AutoFilter {
                    range: rng("A1:A3"),
                    columns: vec![AutoFilterColumn {
                        col: 0,
                        values: Some(vec!["keep".into()]),
                        show_blanks: false,
                    }],
                }),
            },
        )
        .unwrap();
        let sheet = wb.sheet(SheetId(0)).unwrap();
        assert!(!sheet.is_row_hidden(0), "header row must stay visible");
        assert!(!sheet.is_row_hidden(1));
        assert!(sheet.is_row_hidden(2));

        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    fn note(author: &str, text: &str) -> Comment {
        Comment {
            author: author.into(),
            text: text.into(),
        }
    }

    #[test]
    fn set_comment_applies_replaces_deletes_and_inverts() {
        let mut wb = wb_one_sheet();
        let at = r("B2");
        let first = note("Ada", "check this");
        let inv_set = apply(
            &mut wb,
            &Op::SetComment {
                sheet: SheetId(0),
                cell: at,
                comment: Some(first.clone()),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].comment_at(at), Some(&first));

        let second = note("Grace", "done");
        let inv_replace = apply(
            &mut wb,
            &Op::SetComment {
                sheet: SheetId(0),
                cell: at,
                comment: Some(second.clone()),
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].comment_at(at), Some(&second));

        let inv_delete = apply(
            &mut wb,
            &Op::SetComment {
                sheet: SheetId(0),
                cell: at,
                comment: None,
            },
        )
        .unwrap();
        assert!(wb.sheets[0].comments.is_empty());

        apply_ops(&mut wb, &inv_delete.0).unwrap();
        assert_eq!(wb.sheets[0].comment_at(at), Some(&second));
        apply_ops(&mut wb, &inv_replace.0).unwrap();
        assert_eq!(wb.sheets[0].comment_at(at), Some(&first));
        apply_ops(&mut wb, &inv_set.0).unwrap();
        assert!(wb.sheets[0].comments.is_empty());
    }

    #[test]
    fn set_comment_missing_sheet_errors() {
        let mut wb = wb_one_sheet();
        let err = apply(
            &mut wb,
            &Op::SetComment {
                sheet: SheetId(9),
                cell: r("A1"),
                comment: None,
            },
        );
        assert_eq!(err.unwrap_err(), OpError::SheetNotFound(SheetId(9)));
    }

    #[test]
    fn row_and_column_edits_remap_comments_and_undo_restores_them() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_comment(r("A1"), Some(note("Ada", "header")));
        s.set_comment(r("B3"), Some(note("Grace", "moved")));
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].comment_at(r("A1")),
            Some(&note("Ada", "header"))
        );
        assert_eq!(
            wb.sheets[0].comment_at(r("B5")),
            Some(&note("Grace", "moved"))
        );
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);

        let inv = apply(
            &mut wb,
            &Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].comments.len(), 1);
        assert_eq!(
            wb.sheets[0].comment_at(r("A1")),
            Some(&note("Ada", "header"))
        );
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);

        let inv = apply(
            &mut wb,
            &Op::DeleteCols {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(wb.sheets[0].comments.len(), 1);
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);

        let inv = apply(
            &mut wb,
            &Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(
            wb.sheets[0].comment_at(r("C3")),
            Some(&note("Grace", "moved"))
        );
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn removing_a_sheet_restores_its_comments_on_undo() {
        let mut wb = wb_one_sheet();
        wb.sheets.push(Sheet::new("Two"));
        wb.sheet_mut(SheetId(1))
            .unwrap()
            .set_comment(r("C4"), Some(note("Ada", "keep me")));
        let before = wb.clone();

        let inv = apply(&mut wb, &Op::RemoveSheet { index: 1 }).unwrap();
        assert_eq!(wb.sheets.len(), 1);
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn set_hidden_rows_round_trips() {
        let mut wb = wb_one_sheet();
        let inv = apply(
            &mut wb,
            &Op::SetHiddenRows {
                sheet: SheetId(0),
                hidden: vec![1, 3],
            },
        )
        .unwrap();
        assert!(wb.sheets[0].is_row_hidden(1));
        assert!(wb.sheets[0].is_row_hidden(3));
        assert!(!wb.sheets[0].is_row_hidden(2));

        apply_ops(&mut wb, &inv.0).unwrap();
        assert!(wb.sheets[0].hidden_rows.is_empty());
    }

    #[test]
    fn sort_rejects_conflicting_merges_and_allows_uniform_single_row_merges() {
        let sort = Op::SortRange {
            sheet: SheetId(0),
            range: rng("A1:B3"),
            key_col: 0,
            ascending: true,
        };

        for merges in [vec![rng("B3:C3")], vec![rng("A1:A2")], vec![rng("A1:B1")]] {
            let mut wb = wb_one_sheet();
            wb.sheets[0].merges = merges;
            let before = wb.clone();
            let error = apply(&mut wb, &sort).unwrap_err();
            assert!(matches!(error, OpError::MergeConflict { .. }), "{error}");
            assert_eq!(wb, before);
        }

        let mut wb = wb_one_sheet();
        wb.sheets[0].merges = vec![rng("A1:B1"), rng("A2:B2"), rng("A3:B3")];
        apply(&mut wb, &sort).unwrap();
    }

    #[test]
    fn row_edits_remap_hidden_rows_and_auto_filter() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(r("A1"), text_cell("Name"));
        s.set_cell(r("A3"), text_cell("hidden"));
        s.auto_filter = Some(AutoFilter {
            range: rng("A1:A4"),
            columns: vec![AutoFilterColumn {
                col: 0,
                values: None,
                show_blanks: true,
            }],
        });
        s.hidden_rows.insert(2);
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        let sheet = &wb.sheets[0];
        assert_eq!(sheet.auto_filter.as_ref().unwrap().range, rng("A1:A6"));
        assert!(sheet.is_row_hidden(4));
        assert!(!sheet.is_row_hidden(2));
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);

        let inv = apply(
            &mut wb,
            &Op::DeleteRows {
                sheet: SheetId(0),
                at: 2,
                count: 2,
            },
        )
        .unwrap();
        let sheet = &wb.sheets[0];
        assert_eq!(sheet.auto_filter.as_ref().unwrap().range, rng("A1:A2"));
        assert!(sheet.hidden_rows.is_empty());
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn column_edits_remap_auto_filter_range_and_criteria_columns() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(r("B1"), text_cell("Left"));
        s.set_cell(r("C1"), text_cell("Right"));
        s.auto_filter = Some(AutoFilter {
            range: rng("B1:C3"),
            columns: vec![
                AutoFilterColumn {
                    col: 1,
                    values: Some(vec!["x".into()]),
                    show_blanks: false,
                },
                AutoFilterColumn {
                    col: 2,
                    values: None,
                    show_blanks: true,
                },
            ],
        });
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        let filter = wb.sheets[0].auto_filter.as_ref().unwrap();
        assert_eq!(filter.range, rng("C1:D3"));
        assert_eq!(
            filter.columns.iter().map(|c| c.col).collect::<Vec<_>>(),
            [2, 3]
        );
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);

        let inv = apply(
            &mut wb,
            &Op::DeleteCols {
                sheet: SheetId(0),
                at: 1,
                count: 1,
            },
        )
        .unwrap();
        let filter = wb.sheets[0].auto_filter.as_ref().unwrap();
        assert_eq!(filter.range, rng("B1:B3"));
        assert_eq!(
            filter.columns.iter().map(|c| c.col).collect::<Vec<_>>(),
            [1]
        );
        assert!(filter.columns[0].values.is_none());
        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn deleting_the_whole_auto_filter_range_drops_the_filter_and_undo_restores_it() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.auto_filter = Some(AutoFilter {
            range: rng("A2:A3"),
            columns: Vec::new(),
        });
        s.hidden_rows.insert(2);
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::DeleteRows {
                sheet: SheetId(0),
                at: 1,
                count: 2,
            },
        )
        .unwrap();
        assert!(wb.sheets[0].auto_filter.is_none());
        assert!(wb.sheets[0].hidden_rows.is_empty());

        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    fn link(range: &str) -> Hyperlink {
        Hyperlink {
            range: CellRange::parse_a1(range).unwrap(),
            external_target: Some(format!("https://example.com/{range}")),
            location: None,
            tooltip: None,
            display: None,
        }
    }

    fn formula_cell(source: &str) -> Cell {
        Cell {
            formula: Some(source.into()),
            ..Cell::default()
        }
    }

    fn sort_asc(range: &str) -> Op {
        Op::SortRange {
            sheet: SheetId(0),
            range: rng(range),
            key_col: 0,
            ascending: true,
        }
    }

    fn descending_key_column(wb: &mut Workbook) {
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        for (row, value) in [(0, 3.0), (1, 1.0), (2, 2.0)] {
            s.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number { value },
                    ..Cell::default()
                },
            );
        }
    }

    #[test]
    fn sort_range_moves_single_row_hyperlinks_and_undo_restores_them() {
        let mut wb = wb_one_sheet();
        descending_key_column(&mut wb);
        wb.sheets[0]
            .hyperlinks
            .extend([link("A1"), link("A2:B2"), link("C1"), link("A4:B4")]);
        let before = wb.clone();

        let inv = apply(&mut wb, &sort_asc("A1:B3")).unwrap();
        let ranges: Vec<String> = wb.sheets[0]
            .hyperlinks
            .iter()
            .map(|hyperlink| hyperlink.range.to_a1())
            .collect();
        assert_eq!(ranges, vec!["A3", "A1:B1", "C1", "A4:B4"]);

        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn sort_range_rejects_hyperlinks_it_cannot_move_with_a_row() {
        for hyperlink in [link("A1:A2"), link("A2:C2")] {
            let mut wb = wb_one_sheet();
            descending_key_column(&mut wb);
            wb.sheets[0].hyperlinks.push(hyperlink);
            let before = wb.clone();

            let error = apply(&mut wb, &sort_asc("A1:B3")).unwrap_err();
            assert!(
                matches!(error, OpError::HyperlinkConflict { .. }),
                "{error}"
            );
            assert_eq!(wb, before);
        }
    }

    #[test]
    fn sort_range_rewrites_relative_formula_rows_and_undo_restores_the_text() {
        let mut wb = wb_one_sheet();
        descending_key_column(&mut wb);
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(r("B1"), formula_cell("A1*2"));
        s.set_cell(r("B2"), formula_cell("$A$2+A2"));
        s.set_cell(r("B3"), formula_cell("SUM(A3:C3)"));
        s.set_cell(r("D5"), formula_cell("B2+1"));
        let before = wb.clone();

        let inv = apply(&mut wb, &sort_asc("A1:B3")).unwrap();
        assert_eq!(wb.formula(SheetId(0), r("B3")), Some("A3*2"));
        assert_eq!(wb.formula(SheetId(0), r("B1")), Some("$A$2+A1"));
        assert_eq!(wb.formula(SheetId(0), r("B2")), Some("SUM(A2:C2)"));
        assert_eq!(
            wb.formula(SheetId(0), r("D5")),
            Some("B2+1"),
            "formulas outside the sorted rows stay untouched"
        );

        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn move_rows_collapses_underflowing_relative_refs_and_undo_restores_them() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(r("A1"), text_cell("top"));
        s.set_cell(r("A3"), formula_cell("A1+1"));
        let before = wb.clone();

        let inv = apply(
            &mut wb,
            &Op::MoveRows {
                sheet: SheetId(0),
                range: rng("A1:A3"),
                moves: vec![(0, 2), (2, 0)],
            },
        )
        .unwrap();
        assert_eq!(wb.formula(SheetId(0), r("A1")), Some("#REF!+1"));

        apply_ops(&mut wb, &inv.0).unwrap();
        assert_eq!(wb, before);
    }

    #[test]
    fn sort_range_refuses_unparseable_formulas_in_moved_rows() {
        let mut wb = wb_one_sheet();
        descending_key_column(&mut wb);
        wb.sheet_mut(SheetId(0))
            .unwrap()
            .set_cell(r("B1"), formula_cell("SUM("));
        let before = wb.clone();

        let error = apply(&mut wb, &sort_asc("A1:B3")).unwrap_err();
        assert!(
            matches!(error, OpError::FormulaNotRewritable { .. }),
            "{error}"
        );
        assert_eq!(wb, before);
    }

    #[test]
    fn set_auto_filter_preserves_manual_hides_across_filter_changes() {
        let mut wb = wb_one_sheet();
        let s = wb.sheet_mut(SheetId(0)).unwrap();
        s.set_cell(r("A1"), text_cell("Name"));
        s.set_cell(r("A2"), text_cell("keep"));
        s.set_cell(r("A3"), text_cell("drop"));
        s.set_cell(r("A4"), text_cell("keep"));
        let filter_for = |values: &[&str]| {
            Some(AutoFilter {
                range: rng("A1:A4"),
                columns: vec![AutoFilterColumn {
                    col: 0,
                    values: Some(values.iter().map(|v| (*v).to_string()).collect()),
                    show_blanks: false,
                }],
            })
        };
        let state_zero = wb.clone();

        let inv_first = apply(
            &mut wb,
            &Op::SetAutoFilter {
                sheet: SheetId(0),
                filter: filter_for(&["keep"]),
            },
        )
        .unwrap();
        assert!(wb.sheets[0].is_row_hidden(2), "drop row fails the filter");
        let state_one = wb.clone();

        let inv_manual = apply(
            &mut wb,
            &Op::SetHiddenRows {
                sheet: SheetId(0),
                hidden: vec![2, 3],
            },
        )
        .unwrap();
        assert!(
            wb.sheets[0].is_row_hidden(3),
            "manual hide inside the range"
        );
        let state_two = wb.clone();

        let inv_second = apply(
            &mut wb,
            &Op::SetAutoFilter {
                sheet: SheetId(0),
                filter: filter_for(&["keep", "drop"]),
            },
        )
        .unwrap();
        assert!(
            !wb.sheets[0].is_row_hidden(2),
            "the old filter's hidden row passes the new filter"
        );
        assert!(
            wb.sheets[0].is_row_hidden(3),
            "the manual hide survives the filter change"
        );
        let state_three = wb.clone();

        let inv_clear = apply(
            &mut wb,
            &Op::SetAutoFilter {
                sheet: SheetId(0),
                filter: None,
            },
        )
        .unwrap();
        assert!(wb.sheets[0].auto_filter.is_none());
        assert!(
            wb.sheets[0].is_row_hidden(3),
            "clearing the filter keeps the manual hide"
        );
        assert!(!wb.sheets[0].is_row_hidden(2));

        apply_ops(&mut wb, &inv_clear.0).unwrap();
        assert_eq!(wb, state_three);
        apply_ops(&mut wb, &inv_second.0).unwrap();
        assert_eq!(wb, state_two);
        apply_ops(&mut wb, &inv_manual.0).unwrap();
        assert_eq!(wb, state_one);
        apply_ops(&mut wb, &inv_first.0).unwrap();
        assert_eq!(wb, state_zero);
    }

    #[test]
    fn refused_defined_name_rewrites_leave_the_workbook_unchanged() {
        let mut wb = wb_one_sheet();
        wb.sheets[0].set_cell(
            r("B1"),
            Cell {
                formula: Some("A2".into()),
                ..Cell::default()
            },
        );
        wb.defined_names.push(DefinedName {
            name: "Dynamic".into(),
            formula: "OFFSET($A$1,0,0,COUNTA($A:$A),1)".into(),
            local_sheet: Some(SheetId(0)),
            hidden: false,
        });
        let before = wb.clone();

        let error = apply(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap_err();

        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));
        assert_eq!(wb, before);
    }
}
