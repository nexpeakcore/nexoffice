//! wire types for the op log: the serializable `Op` vocabulary, a `CellState`
//! mirror of the model's non-serde `Cell`, provenance tags, and transactions.

use serde::{Deserialize, Serialize};
use xlsx_model::{
    AutoFilter, Cell, CellRange, CellRef, CellValue, ColId, Comment, DefinedName, FreezePane,
    Hyperlink, RowId, SheetId,
};

use crate::formatting::{CapturedFormat, NumberFormatMutation, StylePatch};

/// serializable mirror of `xlsx_model::Cell`, which deliberately has no serde.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellState {
    #[serde(default)]
    pub value: CellValue,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<u32>,
}

impl From<&Cell> for CellState {
    fn from(c: &Cell) -> Self {
        Self {
            value: c.value.clone(),
            formula: c.formula.clone(),
            style: c.style,
        }
    }
}

impl From<Cell> for CellState {
    fn from(c: Cell) -> Self {
        Self {
            value: c.value,
            formula: c.formula,
            style: c.style,
        }
    }
}

impl From<CellState> for Cell {
    fn from(s: CellState) -> Self {
        Cell {
            value: s.value,
            formula: s.formula,
            style: s.style,
        }
    }
}

/// the invertible edit vocabulary. width/height are `Option` so "reset to
/// default" is expressible and invertible.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum Op {
    SetCell {
        sheet: SheetId,
        at: CellRef,
        cell: CellState,
    },
    InsertRows {
        sheet: SheetId,
        at: RowId,
        count: u32,
    },
    DeleteRows {
        sheet: SheetId,
        at: RowId,
        count: u32,
    },
    InsertCols {
        sheet: SheetId,
        at: ColId,
        count: u32,
    },
    DeleteCols {
        sheet: SheetId,
        at: ColId,
        count: u32,
    },
    SetColWidth {
        sheet: SheetId,
        col: ColId,
        width: Option<f64>,
    },
    SetRowHeight {
        sheet: SheetId,
        row: RowId,
        height: Option<f64>,
    },
    SetFreezePane {
        sheet: SheetId,
        pane: Option<FreezePane>,
    },
    SortRange {
        sheet: SheetId,
        /// rows to reorder (inclusive); columns stay fixed — a sort is a row permutation.
        range: CellRange,
        /// absolute sheet column supplying the sort key.
        key_col: ColId,
        ascending: bool,
    },
    /// Absolute row permutation emitted as the inverse of `SortRange`.
    #[doc(hidden)]
    MoveRows {
        sheet: SheetId,
        range: CellRange,
        /// `(from, to)` row pairs, all applied simultaneously.
        moves: Vec<(RowId, RowId)>,
    },
    SetAutoFilter {
        sheet: SheetId,
        /// `None` clears the filter and unhides its rows.
        filter: Option<AutoFilter>,
    },
    /// Absolute hidden-row replacement emitted as the inverse of `SetAutoFilter`.
    #[doc(hidden)]
    SetHiddenRows {
        sheet: SheetId,
        hidden: Vec<RowId>,
        /// the subset of `hidden` attributed to a manual hide rather than the
        /// active filter; restoring it is what makes the undo exact.
        #[serde(default)]
        manual: Vec<RowId>,
    },
    #[doc(hidden)]
    SetHyperlinks {
        sheet: SheetId,
        hyperlinks: Vec<Hyperlink>,
    },
    SetComment {
        sheet: SheetId,
        cell: CellRef,
        /// `None` deletes the comment.
        comment: Option<Comment>,
    },
    MergeCells {
        sheet: SheetId,
        range: CellRange,
    },
    UnmergeCells {
        sheet: SheetId,
        range: CellRange,
    },
    PatchRangeStyle {
        sheet: SheetId,
        range: CellRange,
        patch: StylePatch,
    },
    SetRangeNumberFormat {
        sheet: SheetId,
        range: CellRange,
        format: NumberFormatMutation,
    },
    ApplyRangeFormat {
        sheet: SheetId,
        range: CellRange,
        format: CapturedFormat,
    },
    AddSheet {
        index: usize,
        name: String,
    },
    RemoveSheet {
        index: usize,
    },
    RenameSheet {
        sheet: SheetId,
        name: String,
    },
    /// Absolute replacement emitted as the inverse of structural sheet ops.
    #[doc(hidden)]
    SetDefinedNames {
        defined_names: Vec<DefinedName>,
    },
    #[doc(hidden)]
    RestoreSheet {
        sheet: SheetId,
        name: String,
        formulas: Vec<(SheetId, CellRef, CellState)>,
    },
}

/// who authored a transaction.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum Provenance {
    User,
    Agent { id: String },
    Remote { actor: String },
}

/// an atomic, provenance-tagged batch of ops. `proposed = true` marks a HITL
/// proposal awaiting accept or reject; it is not yet applied.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Transaction {
    pub ops: Vec<Op>,
    pub author: Provenance,
    #[serde(default)]
    pub proposed: bool,
}

impl Transaction {
    pub fn new(ops: Vec<Op>, author: Provenance) -> Self {
        Self {
            ops,
            author,
            proposed: false,
        }
    }

    /// a proposal awaiting review — same payload, not yet committed.
    pub fn proposal(ops: Vec<Op>, author: Provenance) -> Self {
        Self {
            ops,
            author,
            proposed: true,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn set_freeze_pane_wire_shape() {
        let op = Op::SetFreezePane {
            sheet: SheetId(0),
            pane: Some(FreezePane::new(2, 1, CellRef::new(2, 1))),
        };
        let json = serde_json::to_value(&op).unwrap();
        assert_eq!(
            json,
            serde_json::json!({
                "type": "setFreezePane",
                "sheet": 0,
                "pane": {"rows": 2, "cols": 1, "top_left": {"row": 2, "col": 1}}
            })
        );
        let decoded: Op = serde_json::from_value(json).unwrap();
        assert_eq!(decoded, op);

        let unfreeze = Op::SetFreezePane {
            sheet: SheetId(0),
            pane: None,
        };
        let json = serde_json::to_value(&unfreeze).unwrap();
        assert_eq!(
            json,
            serde_json::json!({"type": "setFreezePane", "sheet": 0, "pane": null})
        );
        let decoded: Op = serde_json::from_value(json).unwrap();
        assert_eq!(decoded, unfreeze);
    }

    #[test]
    fn sort_and_filter_wire_shapes() {
        let sort = Op::SortRange {
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:B3").unwrap(),
            key_col: 0,
            ascending: true,
        };
        let json = serde_json::to_value(&sort).unwrap();
        assert_eq!(json["type"], "sortRange");
        let decoded: Op = serde_json::from_value(json).unwrap();
        assert_eq!(decoded, sort);

        let filter = Op::SetAutoFilter {
            sheet: SheetId(0),
            filter: Some(xlsx_model::AutoFilter {
                range: CellRange::parse_a1("A1:B3").unwrap(),
                columns: vec![xlsx_model::AutoFilterColumn {
                    col: 0,
                    values: Some(vec!["x".into()]),
                    show_blanks: true,
                    unsupported: None,
                }],
            }),
        };
        let json = serde_json::to_value(&filter).unwrap();
        assert_eq!(json["type"], "setAutoFilter");
        let decoded: Op = serde_json::from_value(json).unwrap();
        assert_eq!(decoded, filter);

        let clear = Op::SetAutoFilter {
            sheet: SheetId(0),
            filter: None,
        };
        let json = serde_json::to_value(&clear).unwrap();
        assert_eq!(
            json,
            serde_json::json!({"type": "setAutoFilter", "sheet": 0, "filter": null})
        );
        let decoded: Op = serde_json::from_value(json).unwrap();
        assert_eq!(decoded, clear);
    }

    #[test]
    fn set_comment_wire_shape() {
        let op = Op::SetComment {
            sheet: SheetId(0),
            cell: CellRef::new(1, 2),
            comment: Some(Comment {
                author: "Ada".into(),
                text: "check this".into(),
            }),
        };
        let json = serde_json::to_value(&op).unwrap();
        assert_eq!(
            json,
            serde_json::json!({
                "type": "setComment",
                "sheet": 0,
                "cell": {"row": 1, "col": 2},
                "comment": {"author": "Ada", "text": "check this"}
            })
        );
        let decoded: Op = serde_json::from_value(json).unwrap();
        assert_eq!(decoded, op);

        let clear = Op::SetComment {
            sheet: SheetId(0),
            cell: CellRef::new(1, 2),
            comment: None,
        };
        let json = serde_json::to_value(&clear).unwrap();
        assert_eq!(
            json,
            serde_json::json!({
                "type": "setComment",
                "sheet": 0,
                "cell": {"row": 1, "col": 2},
                "comment": null
            })
        );
        let decoded: Op = serde_json::from_value(json).unwrap();
        assert_eq!(decoded, clear);
    }

    #[test]
    fn cell_state_round_trips_through_model_cell() {
        let cell = Cell {
            value: CellValue::Number { value: 3.5 },
            formula: Some("1+2.5".into()),
            style: Some(7),
        };
        let state = CellState::from(&cell);
        assert_eq!(Cell::from(state), cell);
    }
}
