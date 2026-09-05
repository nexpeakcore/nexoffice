//! Functions that produce a rectangle rather than a value.
//!
//! Kept apart from the scalar library on purpose: `BuiltIn` is
//! `fn(&[Expr], &EvalContext) -> CellValue` and 110 functions are written to
//! it, so widening that signature to carry a rectangle nobody else produces
//! would rewrite all of them. A caller that wants a value from one of these
//! takes its top-left, which is what a cell holding the formula shows until
//! something spills it.

use std::cmp::Ordering;
use std::collections::HashMap;
use std::collections::hash_map::Entry;

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, err, evaluate};
use crate::functions::{nth_int, nth_number};
use crate::parser::Expr;

/// Excel's own ceiling on a spilled result. A formula asking for more is
/// `#NUM!` rather than an allocation the size of the sheet.
pub(crate) const MAX_CELLS: usize = 1_048_576;

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

    /// One value as a 1x1 rectangle.
    pub fn scalar(value: CellValue) -> Self {
        Self::new(1, 1, vec![value])
    }

    pub fn at(&self, row: usize, col: usize) -> &CellValue {
        &self.values[row * self.cols + col]
    }

    pub fn values(&self) -> &[CellValue] {
        &self.values
    }

    /// The same rectangle with `f` applied to every cell.
    pub fn map(self, f: impl Fn(CellValue) -> CellValue) -> Self {
        Self {
            rows: self.rows,
            cols: self.cols,
            values: self.values.into_iter().map(f).collect(),
        }
    }

    fn nth(&self, index: usize) -> &CellValue {
        &self.values[index]
    }
}

/// An array function: lazy arguments in, a rectangle out, or the error a cell
/// holding the formula should show.
pub type ArrayBuiltIn = fn(&[Expr], &EvalContext<'_>) -> Result<ArrayValue, CellValue>;

/// Resolve an array function by name, case-insensitively.
pub fn lookup(name: &str) -> Option<ArrayBuiltIn> {
    match name.to_ascii_uppercase().as_str() {
        "FILTER" => Some(filter),
        "SEQUENCE" => Some(sequence),
        "SORT" => Some(sort),
        "SORTBY" => Some(sort_by),
        "UNIQUE" => Some(unique),
        _ => None,
    }
}

/// Which way a function runs through a rectangle: `Rows` takes each row as one
/// entry and works down them, `Cols` each column, and works across.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Axis {
    Rows,
    Cols,
}

impl Axis {
    /// How many entries there are to reorder or keep.
    fn lanes(self, array: &ArrayValue) -> usize {
        match self {
            Axis::Rows => array.rows(),
            Axis::Cols => array.cols(),
        }
    }

    /// How wide one entry is.
    fn width(self, array: &ArrayValue) -> usize {
        match self {
            Axis::Rows => array.cols(),
            Axis::Cols => array.rows(),
        }
    }

    fn at(self, array: &ArrayValue, lane: usize, offset: usize) -> &CellValue {
        match self {
            Axis::Rows => array.at(lane, offset),
            Axis::Cols => array.at(offset, lane),
        }
    }

    /// Rebuild a rectangle from the entries `lanes` names, in that order.
    fn gather(self, array: &ArrayValue, lanes: &[usize]) -> ArrayValue {
        let width = self.width(array);
        let mut values = Vec::with_capacity(lanes.len() * width);
        match self {
            Axis::Rows => {
                for lane in lanes {
                    for offset in 0..width {
                        values.push(array.at(*lane, offset).clone());
                    }
                }
                ArrayValue::new(lanes.len(), width, values)
            }
            Axis::Cols => {
                for offset in 0..width {
                    for lane in lanes {
                        values.push(array.at(offset, *lane).clone());
                    }
                }
                ArrayValue::new(width, lanes.len(), values)
            }
        }
    }
}

/// Read an argument as a rectangle, or a lone value standing in for a 1x1.
fn as_rectangle(arg: &Expr, ctx: &EvalContext<'_>) -> Result<ArrayValue, CellValue> {
    if let Some(area) = crate::eval::as_area(arg, ctx) {
        return area.into_array(ctx).map_err(err);
    }
    match evaluate(arg, ctx) {
        refused @ CellValue::Error { .. } => Err(refused),
        value => Ok(ArrayValue::scalar(value)),
    }
}

/// Which way a key or a mask runs against the rectangle it applies to: a column
/// as tall as the rectangle picks its rows, a row as wide as it picks its
/// columns, and anything else does not line up with it at all.
fn axis_against(vector: &ArrayValue, array: &ArrayValue) -> Option<Axis> {
    if vector.cols() == 1 && vector.rows() == array.rows() {
        Some(Axis::Rows)
    } else if vector.rows() == 1 && vector.cols() == array.cols() {
        Some(Axis::Cols)
    } else {
        None
    }
}

/// Excel's cross-type order, with blanks last whichever way it runs.
///
/// `cmp_values` lets a blank take the other operand's type, which is what
/// `=A1<B1` wants and is not a total order — a blank equal to both `0` and
/// `FALSE` while those differ from each other. `sort_by` is entitled to panic
/// on a comparator like that, so blanks are ranked instead of coerced.
fn ordered(a: &CellValue, b: &CellValue, descending: bool) -> Ordering {
    match (a, b) {
        (CellValue::Empty, CellValue::Empty) => Ordering::Equal,
        (CellValue::Empty, _) => Ordering::Greater,
        (_, CellValue::Empty) => Ordering::Less,
        _ => {
            let ordering = crate::eval::cmp_values(a, b);
            if descending {
                ordering.reverse()
            } else {
                ordering
            }
        }
    }
}

/// A sort direction argument: `1` up, `-1` down, nothing else.
fn direction(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<bool, CellValue> {
    match optional_int(args, ctx, index, 1)? {
        1 => Ok(false),
        -1 => Ok(true),
        _ => Err(err(ErrorValue::Value)),
    }
}

/// Whether an option argument was left for the function to decide.
///
/// Past the end of the list, or a gap between two commas — `SORT(A1:B9,,,TRUE)`
/// sorts by the first row, it does not sort by row zero. That reading is for
/// the option arguments here; a gap where a *value* was wanted is the blank it
/// looks like, which is what `FILTER`'s last argument and `VLOOKUP`'s make of
/// one.
fn left_open(args: &[Expr], index: usize) -> bool {
    matches!(args.get(index), None | Some(Expr::Omitted))
}

fn optional_int(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    default: i64,
) -> Result<i64, CellValue> {
    if left_open(args, index) {
        return Ok(default);
    }
    nth_int(args, ctx, index).map_err(err)
}

fn optional_number(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    default: f64,
) -> Result<f64, CellValue> {
    if left_open(args, index) {
        return Ok(default);
    }
    nth_number(args, ctx, index).map_err(err)
}

fn optional_bool(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    default: bool,
) -> Result<bool, CellValue> {
    if left_open(args, index) {
        return Ok(default);
    }
    crate::eval::to_bool(&evaluate(&args[index], ctx)).map_err(err)
}

/// `SEQUENCE(rows, [columns], [start], [step])`.
fn sequence(args: &[Expr], ctx: &EvalContext<'_>) -> Result<ArrayValue, CellValue> {
    if args.is_empty() || args.len() > 4 {
        return Err(err(ErrorValue::Value));
    }
    let (rows, cols) = (
        optional_int(args, ctx, 0, 1)?,
        optional_int(args, ctx, 1, 1)?,
    );
    if rows < 1 || cols < 1 {
        return Err(err(ErrorValue::Value));
    }
    let cells = (rows as usize)
        .checked_mul(cols as usize)
        .filter(|cells| *cells <= MAX_CELLS)
        .ok_or(err(ErrorValue::Num))?;
    let (start, step) = (
        optional_number(args, ctx, 2, 1.0)?,
        optional_number(args, ctx, 3, 1.0)?,
    );

    let mut values = Vec::with_capacity(cells);
    for index in 0..cells {
        values.push(CellValue::Number {
            value: start + step * index as f64,
        });
    }
    Ok(ArrayValue::new(rows as usize, cols as usize, values))
}

/// `SORT(array, [sort_index], [sort_order], [by_col])`.
fn sort(args: &[Expr], ctx: &EvalContext<'_>) -> Result<ArrayValue, CellValue> {
    if args.is_empty() || args.len() > 4 {
        return Err(err(ErrorValue::Value));
    }
    let array = as_rectangle(&args[0], ctx)?;
    let index = optional_int(args, ctx, 1, 1)?;
    let descending = direction(args, ctx, 2)?;
    let axis = if optional_bool(args, ctx, 3, false)? {
        Axis::Cols
    } else {
        Axis::Rows
    };
    let key = usize::try_from(index)
        .ok()
        .filter(|index| (1..=axis.width(&array)).contains(index))
        .ok_or(err(ErrorValue::Value))?
        - 1;

    let mut lanes: Vec<usize> = (0..axis.lanes(&array)).collect();
    lanes.sort_by(|a, b| {
        ordered(
            axis.at(&array, *a, key),
            axis.at(&array, *b, key),
            descending,
        )
    });
    Ok(axis.gather(&array, &lanes))
}

/// `SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...)`.
fn sort_by(args: &[Expr], ctx: &EvalContext<'_>) -> Result<ArrayValue, CellValue> {
    if args.len() < 2 {
        return Err(err(ErrorValue::Value));
    }
    let array = as_rectangle(&args[0], ctx)?;
    let mut keys: Vec<(ArrayValue, bool)> = Vec::new();
    let mut axis = None;
    let mut index = 1;
    while index < args.len() {
        let by = as_rectangle(&args[index], ctx)?;
        let along = axis_against(&by, &array).ok_or(err(ErrorValue::Value))?;
        if *axis.get_or_insert(along) != along {
            return Err(err(ErrorValue::Value));
        }
        keys.push((by, direction(args, ctx, index + 1)?));
        index += 2;
    }
    let axis = axis.unwrap_or(Axis::Rows);

    let mut lanes: Vec<usize> = (0..axis.lanes(&array)).collect();
    lanes.sort_by(|a, b| {
        keys.iter()
            .map(|(by, descending)| ordered(by.nth(*a), by.nth(*b), *descending))
            .find(|ordering| *ordering != Ordering::Equal)
            .unwrap_or(Ordering::Equal)
    });
    Ok(axis.gather(&array, &lanes))
}

/// `UNIQUE(array, [by_col], [exactly_once])`.
fn unique(args: &[Expr], ctx: &EvalContext<'_>) -> Result<ArrayValue, CellValue> {
    if args.is_empty() || args.len() > 3 {
        return Err(err(ErrorValue::Value));
    }
    let array = as_rectangle(&args[0], ctx)?;
    let axis = if optional_bool(args, ctx, 1, false)? {
        Axis::Cols
    } else {
        Axis::Rows
    };
    let exactly_once = optional_bool(args, ctx, 2, false)?;

    let width = axis.width(&array);
    let mut first_seen: Vec<(usize, usize)> = Vec::new();
    let mut seen: HashMap<Vec<Key>, usize> = HashMap::new();
    for lane in 0..axis.lanes(&array) {
        let key: Vec<Key> = (0..width)
            .map(|offset| Key::of(axis.at(&array, lane, offset)))
            .collect();
        match seen.entry(key) {
            Entry::Occupied(slot) => first_seen[*slot.get()].1 += 1,
            Entry::Vacant(slot) => {
                slot.insert(first_seen.len());
                first_seen.push((lane, 1));
            }
        }
    }

    let lanes: Vec<usize> = first_seen
        .iter()
        .filter(|(_, count)| !exactly_once || *count == 1)
        .map(|(lane, _)| *lane)
        .collect();
    if lanes.is_empty() {
        return Err(err(ErrorValue::Calc));
    }
    Ok(axis.gather(&array, &lanes))
}

/// `FILTER(array, include, [if_empty])`.
fn filter(args: &[Expr], ctx: &EvalContext<'_>) -> Result<ArrayValue, CellValue> {
    if args.len() < 2 || args.len() > 3 {
        return Err(err(ErrorValue::Value));
    }
    let array = as_rectangle(&args[0], ctx)?;
    let include = as_rectangle(&args[1], ctx)?;
    let axis = axis_against(&include, &array).ok_or(err(ErrorValue::Value))?;

    let mut lanes = Vec::new();
    for lane in 0..axis.lanes(&array) {
        if crate::eval::to_bool(include.nth(lane)).map_err(err)? {
            lanes.push(lane);
        }
    }
    if lanes.is_empty() {
        return match args.get(2) {
            Some(if_empty) => as_rectangle(if_empty, ctx),
            None => Err(err(ErrorValue::Calc)),
        };
    }
    Ok(axis.gather(&array, &lanes))
}

/// A value reduced to what makes two entries the same one: text matches
/// case-insensitively and a blank counts as the zero it is shown as, which is
/// what `UNIQUE` treats as a duplicate.
#[derive(PartialEq, Eq, Hash)]
enum Key {
    Number(u64),
    Text(String),
    Bool(bool),
    Error(&'static str),
}

impl Key {
    fn of(value: &CellValue) -> Self {
        match value {
            CellValue::Empty => Key::Number(0.0_f64.to_bits()),
            // `+ 0.0` so that `-0.0` and `0.0` are one entry, not two.
            CellValue::Number { value } => Key::Number((value + 0.0).to_bits()),
            CellValue::Text { value } => Key::Text(value.to_lowercase()),
            CellValue::Bool { value } => Key::Bool(*value),
            CellValue::Error { value } => Key::Error(value.as_str()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{evaluate, evaluate_array, parse_formula};
    use xlsx_model::{Cell, CellRef, DefinedName, Sheet, SheetId, Workbook};

    fn sheet_with(cells: &[(&str, CellValue)]) -> Workbook {
        let mut workbook = Workbook::default();
        let mut sheet = Sheet::new("Sheet1");
        for (address, value) in cells {
            sheet.set_cell(
                CellRef::parse_a1(address).expect("an address"),
                Cell {
                    value: value.clone(),
                    ..Cell::default()
                },
            );
        }
        workbook.sheets.push(sheet);
        workbook
    }

    fn produced_in(workbook: &Workbook, formula: &str) -> Result<ArrayValue, CellValue> {
        let expr = parse_formula(formula).expect("parses");
        let ctx = EvalContext::new(workbook, SheetId(0));
        evaluate_array(&expr, &ctx).expect("an array function")
    }

    fn produced(formula: &str) -> Result<ArrayValue, CellValue> {
        produced_in(&sheet_with(&[]), formula)
    }

    fn shown_in(workbook: &Workbook, formula: &str) -> CellValue {
        let expr = parse_formula(formula).expect("parses");
        evaluate(&expr, &EvalContext::new(workbook, SheetId(0)))
    }

    fn shown(formula: &str) -> CellValue {
        shown_in(&sheet_with(&[]), formula)
    }

    /// `A1:A4` holds 3, 1, 2, 1 and `B1:B4` tags them x, y, x, y; `E1:F3`
    /// names 1, 2, 3.
    fn tagged() -> Workbook {
        sheet_with(&[
            ("A1", n(3.0)),
            ("A2", n(1.0)),
            ("A3", n(2.0)),
            ("A4", n(1.0)),
            ("B1", t("x")),
            ("B2", t("y")),
            ("B3", t("x")),
            ("B4", t("y")),
            ("E1", n(1.0)),
            ("F1", t("mot")),
            ("E2", n(2.0)),
            ("F2", t("hai")),
            ("E3", n(3.0)),
            ("F3", t("ba")),
        ])
    }

    fn e(value: ErrorValue) -> CellValue {
        CellValue::Error { value }
    }

    /// A rectangle as nested rows, so a test can state the answer as a grid.
    fn grid(array: &ArrayValue) -> Vec<Vec<CellValue>> {
        (0..array.rows())
            .map(|row| {
                (0..array.cols())
                    .map(|col| array.at(row, col).clone())
                    .collect()
            })
            .collect()
    }

    fn n(value: f64) -> CellValue {
        CellValue::Number { value }
    }

    fn t(value: &str) -> CellValue {
        CellValue::Text {
            value: value.to_string(),
        }
    }

    fn refused(value: ErrorValue) -> Result<ArrayValue, CellValue> {
        Err(CellValue::Error { value })
    }

    /// `A1:C2` holds two rows of three columns, out of order across.
    fn wide() -> Workbook {
        sheet_with(&[
            ("A1", n(3.0)),
            ("B1", n(1.0)),
            ("C1", n(2.0)),
            ("A2", t("c")),
            ("B2", t("a")),
            ("C2", t("b")),
        ])
    }

    /// `A1:B3` holds three rows, out of order by both of its columns.
    fn table() -> Workbook {
        sheet_with(&[
            ("A1", n(3.0)),
            ("B1", t("ba")),
            ("A2", n(1.0)),
            ("B2", t("mot")),
            ("A3", n(2.0)),
            ("B3", t("hai")),
        ])
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
    fn a_rectangle_is_refused_where_nothing_lifts_it() {
        // CHOOSE's choices are read as values and are not lifted. Saying so
        // beats answering with the first cell, which would be a wrong number
        // rather than a wrong kind — and `#NAME?` would claim the function
        // does not exist, which is not true.
        let refused = CellValue::Error {
            value: ErrorValue::Value,
        };
        assert_eq!(shown("SEQUENCE(3)"), refused);
        assert_eq!(shown("CHOOSE(1,SEQUENCE(3))"), refused);
    }

    #[test]
    fn a_function_that_reads_a_range_reads_a_rectangle_the_same_way() {
        assert_eq!(shown("SUM(SEQUENCE(3))"), n(6.0));
        assert_eq!(shown("COUNTA(SEQUENCE(2,2))"), n(4.0));
        assert_eq!(shown("MAX(SEQUENCE(3,1,10,-1))"), n(10.0));
    }

    #[test]
    fn sorting_moves_whole_rows_and_not_just_the_column_it_reads() {
        let sorted = produced_in(&table(), "SORT(A1:B3)").expect("a rectangle");
        assert_eq!(
            grid(&sorted),
            [[n(1.0), t("mot")], [n(2.0), t("hai")], [n(3.0), t("ba")],]
        );
    }

    #[test]
    fn sorting_reads_the_column_it_is_given_in_the_direction_it_is_given() {
        let by_text = produced_in(&table(), "SORT(A1:B3,2)").expect("a rectangle");
        assert_eq!(
            grid(&by_text),
            [[n(3.0), t("ba")], [n(2.0), t("hai")], [n(1.0), t("mot")],]
        );

        let descending = produced_in(&table(), "SORT(A1:B3,1,-1)").expect("a rectangle");
        assert_eq!(
            grid(&descending),
            [[n(3.0), t("ba")], [n(2.0), t("hai")], [n(1.0), t("mot")],]
        );
    }

    #[test]
    fn by_col_reorders_columns_and_reads_a_row_for_its_key() {
        let sorted = produced_in(&wide(), "SORT(A1:C2,1,1,TRUE)").expect("a rectangle");
        assert_eq!(
            grid(&sorted),
            [[n(1.0), n(2.0), n(3.0)], [t("a"), t("b"), t("c")],]
        );
    }

    #[test]
    fn blanks_sort_last_in_both_directions() {
        // `cmp_values` would let the blank take the other operand's type and
        // land it among the numbers as a 0. Excel puts empty cells after
        // everything, ascending or descending, and a comparator that ranks
        // them is also the only one `sort_by` is allowed to be given.
        let workbook = sheet_with(&[("A1", n(2.0)), ("A3", n(1.0))]);
        let up = produced_in(&workbook, "SORT(A1:A3)").expect("a rectangle");
        assert_eq!(grid(&up), [[n(1.0)], [n(2.0)], [CellValue::Empty]]);

        let down = produced_in(&workbook, "SORT(A1:A3,1,-1)").expect("a rectangle");
        assert_eq!(grid(&down), [[n(2.0)], [n(1.0)], [CellValue::Empty]]);
    }

    #[test]
    fn a_sort_index_off_the_rectangle_or_a_direction_that_is_neither_is_refused() {
        let workbook = table();
        for formula in ["SORT(A1:B3,3)", "SORT(A1:B3,0)", "SORT(A1:B3,1,2)"] {
            assert_eq!(
                produced_in(&workbook, formula),
                refused(ErrorValue::Value),
                "{formula}"
            );
        }
    }

    #[test]
    fn sortby_orders_one_rectangle_by_another() {
        let ordered = produced_in(&table(), "SORTBY(A1:A3,B1:B3)").expect("a rectangle");
        assert_eq!(grid(&ordered), [[n(3.0)], [n(2.0)], [n(1.0)]]);
    }

    #[test]
    fn sortby_falls_through_to_its_later_keys_on_a_tie() {
        let workbook = sheet_with(&[
            ("A1", n(1.0)),
            ("B1", n(2.0)),
            ("A2", n(1.0)),
            ("B2", n(1.0)),
            ("A3", n(0.0)),
            ("B3", n(9.0)),
        ]);
        let ordered =
            produced_in(&workbook, "SORTBY(SEQUENCE(3),A1:A3,1,B1:B3,1)").expect("a rectangle");
        assert_eq!(grid(&ordered), [[n(3.0)], [n(2.0)], [n(1.0)]]);
    }

    #[test]
    fn sortby_refuses_a_key_that_does_not_line_up_with_what_it_sorts() {
        assert_eq!(
            produced_in(&table(), "SORTBY(A1:A3,B1:B2)"),
            refused(ErrorValue::Value)
        );
    }

    #[test]
    fn unique_keeps_the_first_spelling_of_each_entry() {
        let workbook = sheet_with(&[("A1", t("Ha")), ("A2", t("ha")), ("A3", t("Noi"))]);
        let kept = produced_in(&workbook, "UNIQUE(A1:A3)").expect("a rectangle");
        assert_eq!(grid(&kept), [[t("Ha")], [t("Noi")]]);
    }

    #[test]
    fn unique_compares_whole_rows_not_single_cells() {
        let workbook = sheet_with(&[
            ("A1", n(1.0)),
            ("B1", n(2.0)),
            ("A2", n(1.0)),
            ("B2", n(2.0)),
            ("A3", n(1.0)),
            ("B3", n(3.0)),
        ]);
        let kept = produced_in(&workbook, "UNIQUE(A1:B3)").expect("a rectangle");
        assert_eq!(grid(&kept), [[n(1.0), n(2.0)], [n(1.0), n(3.0)]]);
    }

    #[test]
    fn unique_can_keep_only_what_appears_exactly_once() {
        let workbook = sheet_with(&[("A1", n(1.0)), ("A2", n(1.0)), ("A3", n(2.0))]);
        let once = produced_in(&workbook, "UNIQUE(A1:A3,FALSE,TRUE)").expect("a rectangle");
        assert_eq!(grid(&once), [[n(2.0)]]);
    }

    #[test]
    fn a_result_with_nothing_in_it_says_so() {
        let workbook = sheet_with(&[("A1", n(1.0)), ("A2", n(1.0)), ("B1", n(0.0))]);
        assert_eq!(
            produced_in(&workbook, "UNIQUE(A1:A2,FALSE,TRUE)"),
            refused(ErrorValue::Calc)
        );
        assert_eq!(
            produced_in(&workbook, "FILTER(A1:A2,B1:C1)"),
            refused(ErrorValue::Value),
            "a mask matching neither the height nor the width does not line up"
        );
    }

    #[test]
    fn a_single_cell_mask_answers_for_a_whole_column() {
        // A 1x1 mask is not as tall as the column, but it is as wide, so it
        // keeps or drops the column whole. A mask has to match one of the two
        // dimensions and this one matches only that one.
        let workbook = sheet_with(&[("A1", n(1.0)), ("A2", n(2.0)), ("D1", n(1.0))]);
        let kept = produced_in(&workbook, "FILTER(A1:A2,D1:D1)").expect("a rectangle");
        assert_eq!(grid(&kept), [[n(1.0)], [n(2.0)]]);
    }

    #[test]
    fn filter_keeps_the_rows_its_mask_marks() {
        let workbook = sheet_with(&[
            ("A1", n(3.0)),
            ("B1", t("ba")),
            ("A2", n(1.0)),
            ("B2", t("mot")),
            ("A3", n(2.0)),
            ("B3", t("hai")),
            ("D1", CellValue::Bool { value: true }),
            ("D2", CellValue::Bool { value: false }),
            ("D3", CellValue::Bool { value: true }),
        ]);
        let kept = produced_in(&workbook, "FILTER(A1:B3,D1:D3)").expect("a rectangle");
        assert_eq!(grid(&kept), [[n(3.0), t("ba")], [n(2.0), t("hai")]]);
    }

    #[test]
    fn filter_answers_with_what_it_was_told_to_say_when_nothing_passes() {
        let workbook = sheet_with(&[
            ("A1", n(1.0)),
            ("A2", n(2.0)),
            ("D1", n(0.0)),
            ("D2", n(0.0)),
        ]);
        assert_eq!(
            produced_in(&workbook, "FILTER(A1:A2,D1:D2)"),
            refused(ErrorValue::Calc)
        );
        let told = produced_in(&workbook, "FILTER(A1:A2,D1:D2,\"khong co\")").expect("a rectangle");
        assert_eq!(grid(&told), [[t("khong co")]]);
    }

    #[test]
    fn filter_hands_back_an_error_standing_in_its_mask() {
        let workbook = sheet_with(&[
            ("A1", n(1.0)),
            ("A2", n(2.0)),
            (
                "D1",
                CellValue::Error {
                    value: ErrorValue::NA,
                },
            ),
            ("D2", n(1.0)),
        ]);
        assert_eq!(
            produced_in(&workbook, "FILTER(A1:A2,D1:D2)"),
            refused(ErrorValue::NA)
        );
        assert_eq!(
            produced_in(
                &sheet_with(&[("A1", n(1.0)), ("D1", t("co"))]),
                "FILTER(A1:A1,D1:D1)"
            ),
            refused(ErrorValue::Value)
        );
    }

    #[test]
    fn an_option_left_open_is_the_one_the_function_would_have_chosen() {
        // The spelling all of this is for: sorting by column is the fourth
        // argument and nobody writes the two in front of it. A gap has to mean
        // "sort by the first row", not "sort by row zero".
        let sorted = produced_in(&wide(), "SORT(A1:C2,,,TRUE)").expect("a rectangle");
        assert_eq!(
            grid(&sorted),
            [[n(1.0), n(2.0), n(3.0)], [t("a"), t("b"), t("c")]]
        );

        let workbook = sheet_with(&[("A1", n(1.0)), ("A2", n(1.0)), ("A3", n(2.0))]);
        let once = produced_in(&workbook, "UNIQUE(A1:A3,,TRUE)").expect("a rectangle");
        assert_eq!(grid(&once), [[n(2.0)]]);
    }

    #[test]
    fn a_comparison_across_a_column_is_the_mask_a_filter_reads() {
        // The way FILTER is written: the condition is an operator applied to
        // a range, not a column of TRUE/FALSE prepared in advance.
        let kept = produced_in(&tagged(), "FILTER(A1:A4,B1:B4=\"x\")").expect("a rectangle");
        assert_eq!(grid(&kept), [[n(3.0)], [n(2.0)]]);

        // Two conditions are multiplied, which is how AND is spelled across
        // a rectangle.
        let both =
            produced_in(&tagged(), "FILTER(A1:A4,(B1:B4=\"x\")*(A1:A4>2))").expect("a rectangle");
        assert_eq!(grid(&both), [[n(3.0)]]);
    }

    #[test]
    fn a_function_that_reads_a_range_reads_what_a_rectangle_produced() {
        let workbook = tagged();
        assert_eq!(
            shown_in(&workbook, "SUM(FILTER(A1:A4,B1:B4=\"x\"))"),
            n(5.0)
        );
        assert_eq!(shown_in(&workbook, "COUNTA(UNIQUE(A1:A4))"), n(3.0));
        assert_eq!(
            shown_in(&workbook, "ROWS(FILTER(A1:A4,B1:B4=\"x\"))"),
            n(2.0)
        );
        assert_eq!(
            shown_in(&workbook, "SUMPRODUCT((B1:B4=\"x\")*A1:A4)"),
            n(5.0)
        );
        assert_eq!(shown_in(&workbook, "SUM(A1:A4*2)"), n(14.0));
        assert_eq!(
            shown_in(&workbook, "TEXTJOIN(\",\",TRUE,UNIQUE(B1:B4))"),
            t("x,y")
        );
        assert_eq!(shown_in(&workbook, "INDEX(SORT(A1:A4),1)"), n(1.0));
        assert_eq!(shown_in(&workbook, "COUNTIF(UNIQUE(A1:A4),\">1\")"), n(2.0));
    }

    #[test]
    fn an_operator_runs_across_a_rectangle_cell_by_cell() {
        let doubled = produced_in(&tagged(), "A1:A4*2").expect("a rectangle");
        assert_eq!(grid(&doubled), [[n(6.0)], [n(2.0)], [n(4.0)], [n(2.0)]]);
        let negated = produced("-SEQUENCE(2)").expect("a rectangle");
        assert_eq!(grid(&negated), [[n(-1.0)], [n(-2.0)]]);
        let percent = produced("SEQUENCE(2)%").expect("a rectangle");
        assert_eq!(grid(&percent), [[n(0.01)], [n(0.02)]]);
    }

    #[test]
    fn sides_that_disagree_pad_with_na_and_a_column_against_a_row_is_every_pairing() {
        let na = CellValue::Error {
            value: ErrorValue::NA,
        };
        let padded = produced("SEQUENCE(3)+SEQUENCE(2)").expect("a rectangle");
        assert_eq!(grid(&padded), [[n(2.0)], [n(4.0)], [na]]);

        let table = produced("SEQUENCE(3)*SEQUENCE(1,2)").expect("a rectangle");
        assert_eq!(
            grid(&table),
            [[n(1.0), n(2.0)], [n(2.0), n(4.0)], [n(3.0), n(6.0)]]
        );
        assert_eq!(
            produced("SEQUENCE(1024)*SEQUENCE(1,1025)"),
            refused(ErrorValue::Num),
            "every pairing of those is more cells than a sheet has"
        );
    }

    #[test]
    fn a_lone_cell_is_a_value_and_does_not_make_a_rectangle() {
        // `=A1*2` must not spill: a formula that spills is saved as an array
        // formula, and that would turn every plain formula in a workbook
        // into one.
        let workbook = tagged();
        let ctx = EvalContext::new(&workbook, SheetId(0));
        for formula in ["A1*2", "A1", "-A1", "A1&B1", "SUM(A1:A4)*2"] {
            let expr = parse_formula(formula).expect("parses");
            assert!(evaluate_array(&expr, &ctx).is_none(), "{formula}");
        }
    }

    #[test]
    fn a_producer_that_fails_is_read_as_the_error_it_produced() {
        let workbook = tagged();
        let calc = CellValue::Error {
            value: ErrorValue::Calc,
        };
        assert_eq!(shown_in(&workbook, "SUM(FILTER(A1:A4,B1:B4=\"z\"))"), calc);
        assert_eq!(
            shown_in(&workbook, "IFERROR(SUM(FILTER(A1:A4,B1:B4=\"z\")),0)"),
            n(0.0)
        );
        // Excel's answer too: COUNTA counts an error as something, and the
        // failed filter is one error. The common idiom is ROWS, above.
        assert_eq!(
            shown_in(&workbook, "COUNTA(FILTER(A1:A4,B1:B4=\"z\"))"),
            n(1.0)
        );
    }

    #[test]
    fn a_name_is_whatever_it_is_defined_as() {
        let mut workbook = tagged();
        workbook.defined_names.push(DefinedName {
            name: "Scores".into(),
            formula: "Sheet1!A1:A4".into(),
            local_sheet: None,
            hidden: false,
        });
        workbook.defined_names.push(DefinedName {
            name: "Top".into(),
            formula: "Sheet1!A1".into(),
            local_sheet: None,
            hidden: false,
        });
        let ctx = EvalContext::new(&workbook, SheetId(0));
        let doubled = evaluate_array(&parse_formula("Scores*2").unwrap(), &ctx)
            .expect("a rectangle")
            .expect("produced");
        assert_eq!(grid(&doubled), [[n(6.0)], [n(2.0)], [n(4.0)], [n(2.0)]]);
        assert!(evaluate_array(&parse_formula("Top*2").unwrap(), &ctx).is_none());
    }

    #[test]
    fn a_function_that_wants_a_reference_does_not_take_a_rectangle_for_one() {
        assert_eq!(
            shown_in(&tagged(), "ROW(SORT(A1:A4))"),
            CellValue::Error {
                value: ErrorValue::Value
            }
        );
    }

    #[test]
    fn a_function_given_a_rectangle_where_it_takes_a_value_runs_once_per_cell() {
        let workbook = tagged();
        let lengths = produced_in(&workbook, "LEN(REPT(B1:B4,A1:A4))").expect("a rectangle");
        assert_eq!(grid(&lengths), [[n(3.0)], [n(1.0)], [n(2.0)], [n(1.0)]]);
        let rounded = produced_in(&workbook, "ROUND(A1:A4/3,1)").expect("a rectangle");
        assert_eq!(grid(&rounded), [[n(1.0)], [n(0.3)], [n(0.7)], [n(0.3)]]);
        let shown = produced_in(&workbook, "TEXT(A1:A4,\"0.0\")").expect("a rectangle");
        assert_eq!(
            grid(&shown),
            [[t("3.0")], [t("1.0")], [t("2.0")], [t("1.0")]]
        );
    }

    #[test]
    fn the_conditions_people_write_with_a_scalar_function_are_masks() {
        let workbook = tagged();
        let found = produced_in(&workbook, "ISNUMBER(SEARCH(\"x\",B1:B4))").expect("a rectangle");
        let (yes, no) = (
            CellValue::Bool { value: true },
            CellValue::Bool { value: false },
        );
        assert_eq!(grid(&found), [[yes.clone()], [no.clone()], [yes], [no]]);

        let kept = produced_in(&workbook, "FILTER(A1:A4,ISNUMBER(SEARCH(\"x\",B1:B4)))")
            .expect("a rectangle");
        assert_eq!(grid(&kept), [[n(3.0)], [n(2.0)]]);
        let initial =
            produced_in(&workbook, "FILTER(A1:A4,LEFT(B1:B4,1)=\"y\")").expect("a rectangle");
        assert_eq!(grid(&initial), [[n(1.0)], [n(1.0)]]);
    }

    #[test]
    fn if_runs_once_per_cell_and_still_takes_only_the_branch_it_needs() {
        let workbook = tagged();
        let sized = produced_in(&workbook, "IF(A1:A4>2,\"big\",\"small\")").expect("a rectangle");
        assert_eq!(
            grid(&sized),
            [[t("big")], [t("small")], [t("small")], [t("small")]]
        );

        let lazy = produced_in(&workbook, "IF(A1:A4>2,1/0,\"b\")").expect("a rectangle");
        assert_eq!(
            grid(&lazy),
            [[e(ErrorValue::Div0)], [t("b")], [t("b")], [t("b")]]
        );

        // The idiom that predates SUMIF, and what people still write.
        assert_eq!(shown_in(&workbook, "SUM(IF(A1:A4>1,A1:A4,0))"), n(5.0));
    }

    #[test]
    fn a_value_parameter_lifts_and_a_range_parameter_beside_it_does_not() {
        let workbook = tagged();
        let counts = produced_in(&workbook, "COUNTIF(B1:B4,UNIQUE(B1:B4))").expect("a rectangle");
        assert_eq!(grid(&counts), [[n(2.0)], [n(2.0)]]);
        let top = produced_in(&workbook, "LARGE(A1:A4,SEQUENCE(2))").expect("a rectangle");
        assert_eq!(grid(&top), [[n(3.0)], [n(2.0)]]);
        let ranks = produced_in(&workbook, "RANK(A1:A4,A1:A4)").expect("a rectangle");
        assert_eq!(grid(&ranks), [[n(1.0)], [n(3.0)], [n(2.0)], [n(3.0)]]);
        let named = produced_in(&workbook, "VLOOKUP(A1:A4,E1:F3,2,)").expect("a rectangle");
        assert_eq!(
            grid(&named),
            [[t("ba")], [t("mot")], [t("hai")], [t("mot")]]
        );
        let picked = produced_in(&workbook, "INDEX(E1:F3,SEQUENCE(2),2)").expect("a rectangle");
        assert_eq!(grid(&picked), [[t("mot")], [t("hai")]]);
        let over = produced_in(&workbook, "SUMIF(A1:A4,\">\"&SEQUENCE(2))").expect("a rectangle");
        assert_eq!(grid(&over), [[n(5.0)], [n(3.0)]]);
    }

    #[test]
    fn a_lifted_function_nests_in_another() {
        let workbook = tagged();
        let safe =
            produced_in(&workbook, "IFERROR(VLOOKUP(A1:A5,E1:F3,2,),\"-\")").expect("a rectangle");
        assert_eq!(
            grid(&safe),
            [[t("ba")], [t("mot")], [t("hai")], [t("mot")], [t("-")]]
        );
        assert_eq!(shown_in(&workbook, "SUM(LEN(REPT(B1:B4,A1:A4)))"), n(7.0));
    }

    #[test]
    fn rectangles_in_several_value_parameters_line_up_like_an_operators_sides() {
        let chosen = produced("IF(SEQUENCE(3)>1,SEQUENCE(2),\"c\")").expect("a rectangle");
        assert_eq!(grid(&chosen), [[t("c")], [n(2.0)], [e(ErrorValue::NA)]]);
        assert_eq!(
            produced("ROUND(SEQUENCE(1024),SEQUENCE(1,1025))"),
            refused(ErrorValue::Num)
        );
    }

    #[test]
    fn a_blank_cell_is_handed_over_as_a_blank() {
        let workbook = tagged();
        let lengths = produced_in(&workbook, "LEN(A4:A5)").expect("a rectangle");
        assert_eq!(grid(&lengths), [[n(1.0)], [n(0.0)]]);
        let blank = produced_in(&workbook, "ISBLANK(A4:A5)").expect("a rectangle");
        assert_eq!(
            grid(&blank),
            [
                [CellValue::Bool { value: false }],
                [CellValue::Bool { value: true }]
            ]
        );
    }

    #[test]
    fn a_formula_with_no_rectangle_in_a_value_parameter_is_not_lifted() {
        let workbook = tagged();
        let ctx = EvalContext::new(&workbook, SheetId(0));
        for formula in [
            "ROUND(A1,0)",
            "IF(A1>2,\"a\",\"b\")",
            "VLOOKUP(A1,E1:F3,2,)",
            "SUM(A1:A4)",
            "AND(A1:A4>0)",
            "ROWS(A1:A4)",
            "SUM(LEN(B1:B4))",
            "COUNTIF(B1:B4,\"x\")",
        ] {
            let expr = parse_formula(formula).expect("parses");
            assert!(evaluate_array(&expr, &ctx).is_none(), "{formula}");
        }
    }

    #[test]
    fn one_rectangle_can_be_read_out_of_another() {
        let workbook = sheet_with(&[
            ("A1", n(3.0)),
            ("A2", n(1.0)),
            ("A3", n(3.0)),
            ("A4", n(2.0)),
        ]);
        let sorted = produced_in(&workbook, "SORT(UNIQUE(A1:A4))").expect("a rectangle");
        assert_eq!(grid(&sorted), [[n(1.0)], [n(2.0)], [n(3.0)]]);
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
