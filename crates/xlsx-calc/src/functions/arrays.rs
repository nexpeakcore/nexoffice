//! Functions that produce a rectangle rather than a value.
//!
//! Kept apart from the scalar library on purpose: `BuiltIn` is
//! `fn(&[Expr], &EvalContext) -> CellValue` and 110 functions are written to
//! it, so widening that signature to carry a rectangle nobody else produces
//! would rewrite all of them. A caller that wants a value from one of these
//! takes its top-left, which is what a cell holding the formula shows until
//! something spills it.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::EvalContext;
use crate::functions::nth_int;
use crate::parser::Expr;

/// Excel's own ceiling on a spilled result. A formula asking for more is
/// `#NUM!` rather than an allocation the size of the sheet.
const MAX_CELLS: usize = 1_048_576;

/// A rectangle of values, in row-major order.
#[derive(Debug, Clone, PartialEq)]
pub struct ArrayValue {
    rows: usize,
    cols: usize,
    values: Vec<CellValue>,
}

impl ArrayValue {
    /// `values` must hold exactly `rows * cols` entries, row-major.
    pub fn new(rows: usize, cols: usize, values: Vec<CellValue>) -> Self {
        debug_assert_eq!(rows * cols, values.len());
        Self { rows, cols, values }
    }

    pub fn rows(&self) -> usize {
        self.rows
    }

    pub fn cols(&self) -> usize {
        self.cols
    }

    pub fn at(&self, row: usize, col: usize) -> &CellValue {
        &self.values[row * self.cols + col]
    }
}

/// An array function: lazy arguments in, a rectangle out, or the error a cell
/// holding the formula should show.
pub type ArrayBuiltIn = fn(&[Expr], &EvalContext<'_>) -> Result<ArrayValue, CellValue>;

/// Resolve an array function by name, case-insensitively.
pub fn lookup(name: &str) -> Option<ArrayBuiltIn> {
    match name.to_ascii_uppercase().as_str() {
        "SEQUENCE" => Some(sequence),
        _ => None,
    }
}

/// `SEQUENCE(rows, [columns], [start], [step])`.
fn sequence(args: &[Expr], ctx: &EvalContext<'_>) -> Result<ArrayValue, CellValue> {
    if args.is_empty() || args.len() > 4 {
        return Err(CellValue::Error {
            value: ErrorValue::Value,
        });
    }
    let count = |index: usize, default: i64| -> Result<i64, CellValue> {
        if index >= args.len() {
            return Ok(default);
        }
        nth_int(args, ctx, index).map_err(|value| CellValue::Error { value })
    };
    let (rows, cols) = (count(0, 1)?, count(1, 1)?);
    if rows < 1 || cols < 1 {
        return Err(CellValue::Error {
            value: ErrorValue::Value,
        });
    }
    let cells = (rows as usize)
        .checked_mul(cols as usize)
        .filter(|cells| *cells <= MAX_CELLS)
        .ok_or(CellValue::Error {
            value: ErrorValue::Num,
        })?;

    let number = |index: usize, default: f64| -> Result<f64, CellValue> {
        if index >= args.len() {
            return Ok(default);
        }
        crate::functions::nth_number(args, ctx, index).map_err(|value| CellValue::Error { value })
    };
    let (start, step) = (number(2, 1.0)?, number(3, 1.0)?);

    let mut values = Vec::with_capacity(cells);
    for index in 0..cells {
        values.push(CellValue::Number {
            value: start + step * index as f64,
        });
    }
    Ok(ArrayValue::new(rows as usize, cols as usize, values))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{evaluate, evaluate_array, parse_formula};
    use xlsx_model::{Sheet, SheetId, Workbook};

    fn produced(formula: &str) -> Result<ArrayValue, CellValue> {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Sheet1"));
        let expr = parse_formula(formula).expect("parses");
        let ctx = EvalContext::new(&workbook, SheetId(0));
        evaluate_array(&expr, &ctx).expect("an array function")
    }

    fn shown(formula: &str) -> CellValue {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Sheet1"));
        let expr = parse_formula(formula).expect("parses");
        evaluate(&expr, &EvalContext::new(&workbook, SheetId(0)))
    }

    #[test]
    fn counts_down_a_column_by_default() {
        let array = produced("SEQUENCE(3)").expect("a rectangle");
        assert_eq!((array.rows(), array.cols()), (3, 1));
        assert_eq!(*array.at(2, 0), CellValue::Number { value: 3.0 });
    }

    #[test]
    fn fills_a_rectangle_across_before_down() {
        let array = produced("SEQUENCE(2,3)").expect("a rectangle");
        assert_eq!((array.rows(), array.cols()), (2, 3));
        assert_eq!(*array.at(0, 2), CellValue::Number { value: 3.0 });
        assert_eq!(*array.at(1, 0), CellValue::Number { value: 4.0 });
    }

    #[test]
    fn starts_and_steps_where_it_is_told() {
        let array = produced("SEQUENCE(3,1,10,-2.5)").expect("a rectangle");
        assert_eq!(*array.at(0, 0), CellValue::Number { value: 10.0 });
        assert_eq!(*array.at(2, 0), CellValue::Number { value: 5.0 });
    }

    #[test]
    fn refuses_an_empty_or_impossible_rectangle() {
        for formula in ["SEQUENCE(0)", "SEQUENCE(2,0)", "SEQUENCE(-1)"] {
            assert_eq!(
                produced(formula),
                Err(CellValue::Error {
                    value: ErrorValue::Value
                }),
                "{formula}"
            );
        }
        assert_eq!(
            produced("SEQUENCE(1048576,2)"),
            Err(CellValue::Error {
                value: ErrorValue::Num
            })
        );
    }

    #[test]
    fn a_rectangle_is_refused_where_a_value_was_wanted() {
        // Until something can spill it, a rectangle has nowhere to go. Saying
        // so beats answering with its first cell, which would make
        // `SUM(SEQUENCE(3))` quietly 1 — and `#NAME?` would claim the function
        // does not exist, which is no longer true.
        let refused = CellValue::Error {
            value: ErrorValue::Value,
        };
        assert_eq!(shown("SEQUENCE(3)"), refused);
        assert_eq!(shown("SUM(SEQUENCE(3))"), refused);
    }

    #[test]
    fn only_a_whole_formula_offers_a_rectangle() {
        let mut workbook = Workbook::default();
        workbook.sheets.push(Sheet::new("Sheet1"));
        let ctx = EvalContext::new(&workbook, SheetId(0));
        let nested = parse_formula("SUM(SEQUENCE(3))").expect("parses");
        assert!(evaluate_array(&nested, &ctx).is_none());
    }
}
