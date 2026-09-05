//! builtin function library: `lookup` maps a case-insensitive name to a builtin.
//! builtins receive arguments unevaluated so control-flow can skip branches.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, as_area, err, evaluate, num, parse_num, to_number};
use crate::parser::Expr;

pub mod arrays;
pub mod criteria;
pub mod datetime;
pub mod info;
pub mod logical;
pub mod lookups;
pub mod math;
pub mod stats;
pub mod text;

/// a builtin: lazy arguments in, one value out.
pub type BuiltIn = fn(&[Expr], &EvalContext<'_>) -> CellValue;

/// Which of a function's parameters take one value, so that a rectangle
/// given for one runs the function once per cell, and which take a range,
/// which is read whole.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Params {
    /// Every parameter takes one value: `LEN`, `ROUND`, `IF`.
    Values,
    /// Every parameter takes a range: `SUM`, `AND`, `ROWS`.
    Ranges,
    /// These positions take one value and the rest take a range.
    ValuesAt(&'static [usize]),
    /// A criterion at every second position from here on, each after the
    /// range it applies to: `COUNTIFS(range, crit, range, crit)`.
    CriteriaFrom(usize),
}

impl Params {
    pub(crate) fn takes_a_value(self, index: usize) -> bool {
        match self {
            Params::Values => true,
            Params::Ranges => false,
            Params::ValuesAt(positions) => positions.contains(&index),
            Params::CriteriaFrom(first) => index >= first && (index - first).is_multiple_of(2),
        }
    }
}

/// The parameter shape of a builtin. Anything not named here is read as
/// ranges, which is what every function did before rectangles could be
/// lifted, so an omission costs a lift and never a wrong answer.
pub(crate) fn params(name: &str) -> Params {
    match name.to_ascii_uppercase().as_str() {
        "ABS" | "SIGN" | "ROUND" | "ROUNDUP" | "ROUNDDOWN" | "MROUND" | "CEILING" | "FLOOR"
        | "INT" | "TRUNC" | "MOD" | "POWER" | "SQRT" | "EXP" | "LN" | "LOG" | "LOG10" => {
            Params::Values
        }
        "LEN" | "LEFT" | "RIGHT" | "MID" | "FIND" | "SEARCH" | "SUBSTITUTE" | "REPLACE"
        | "TRIM" | "UPPER" | "LOWER" | "PROPER" | "CLEAN" | "REPT" | "EXACT" | "T" | "CHAR"
        | "CODE" | "VALUE" | "NUMBERVALUE" | "TEXT" => Params::Values,
        "DATE" | "YEAR" | "MONTH" | "DAY" | "WEEKDAY" | "EDATE" | "EOMONTH" | "HOUR" | "MINUTE"
        | "SECOND" | "TIME" | "DATEDIF" => Params::Values,
        "IF" | "IFERROR" | "IFNA" | "IFS" | "SWITCH" | "NOT" => Params::Values,
        "ISBLANK" | "ISNUMBER" | "ISTEXT" | "ISLOGICAL" | "ISERROR" | "ISERR" | "ISNA" | "N" => {
            Params::Values
        }
        "SUMIF" | "COUNTIF" | "AVERAGEIF" | "LARGE" | "SMALL" => Params::ValuesAt(&[1]),
        "COUNTIFS" => Params::CriteriaFrom(1),
        "SUMIFS" | "AVERAGEIFS" => Params::CriteriaFrom(2),
        "RANK" | "RANK.EQ" | "MATCH" => Params::ValuesAt(&[0, 2]),
        "VLOOKUP" | "HLOOKUP" => Params::ValuesAt(&[0, 2, 3]),
        "INDEX" => Params::ValuesAt(&[1, 2]),
        "XLOOKUP" => Params::ValuesAt(&[0, 3, 4, 5]),
        "CHOOSE" => Params::ValuesAt(&[0]),
        "TEXTJOIN" => Params::ValuesAt(&[0, 1]),
        _ => Params::Ranges,
    }
}

/// resolve a function name (case-insensitive) to its implementation; aliases
/// map to the same function.
pub fn lookup(name: &str) -> Option<BuiltIn> {
    let upper = name.to_ascii_uppercase();
    Some(match upper.as_str() {
        "SUM" => math::sum,
        "SUMIF" => math::sumif,
        "SUMIFS" => math::sumifs,
        "SUMPRODUCT" => math::sumproduct,
        "PRODUCT" => math::product,
        "ABS" => math::abs,
        "SIGN" => math::sign,
        "ROUND" => math::round,
        "ROUNDUP" => math::roundup,
        "ROUNDDOWN" => math::rounddown,
        "MROUND" => math::mround,
        "CEILING" => math::ceiling,
        "FLOOR" => math::floor,
        "INT" => math::int,
        "TRUNC" => math::trunc,
        "MOD" => math::mod_,
        "POWER" => math::power,
        "SQRT" => math::sqrt,
        "EXP" => math::exp,
        "LN" => math::ln,
        "LOG" => math::log,
        "LOG10" => math::log10,
        "PI" => math::pi,
        "AVERAGE" => stats::average,
        "COUNT" => stats::count,
        "COUNTA" => stats::counta,
        "COUNTBLANK" => stats::countblank,
        "COUNTIF" => stats::countif,
        "COUNTIFS" => stats::countifs,
        "AVERAGEIF" => stats::averageif,
        "AVERAGEIFS" => stats::averageifs,
        "MIN" => stats::min,
        "MAX" => stats::max,
        "MEDIAN" => stats::median,
        "MODE" | "MODE.SNGL" => stats::mode,
        "STDEV" | "STDEV.S" => stats::stdev_s,
        "STDEVP" | "STDEV.P" => stats::stdev_p,
        "VAR" | "VAR.S" => stats::var_s,
        "VARP" | "VAR.P" => stats::var_p,
        "LARGE" => stats::large,
        "SMALL" => stats::small,
        "RANK" | "RANK.EQ" => stats::rank,
        "LEN" => text::len,
        "LEFT" => text::left,
        "RIGHT" => text::right,
        "MID" => text::mid,
        "FIND" => text::find,
        "SEARCH" => text::search,
        "SUBSTITUTE" => text::substitute,
        "REPLACE" => text::replace,
        "TRIM" => text::trim,
        "UPPER" => text::upper,
        "LOWER" => text::lower,
        "PROPER" => text::proper,
        "CLEAN" => text::clean,
        "REPT" => text::rept,
        "EXACT" => text::exact,
        "T" => text::t,
        "CHAR" => text::char_,
        "CODE" => text::code,
        "VALUE" => text::value,
        "NUMBERVALUE" => text::numbervalue,
        "TEXT" => text::text_fn,
        "TEXTJOIN" => text::textjoin,
        "CONCATENATE" | "CONCAT" => text::concat,
        "DATE" => datetime::date,
        "YEAR" => datetime::year,
        "MONTH" => datetime::month,
        "DAY" => datetime::day,
        "WEEKDAY" => datetime::weekday,
        "EDATE" => datetime::edate,
        "EOMONTH" => datetime::eomonth,
        "TODAY" => datetime::today,
        "NOW" => datetime::now,
        "HOUR" => datetime::hour,
        "MINUTE" => datetime::minute,
        "SECOND" => datetime::second,
        "TIME" => datetime::time,
        "DATEDIF" => datetime::datedif,
        "IF" => logical::if_,
        "IFERROR" => logical::iferror,
        "IFNA" => logical::ifna,
        "IFS" => logical::ifs,
        "SWITCH" => logical::switch,
        "AND" => logical::and,
        "OR" => logical::or,
        "NOT" => logical::not,
        "XOR" => logical::xor,
        "VLOOKUP" => lookups::vlookup,
        "HLOOKUP" => lookups::hlookup,
        "INDEX" => lookups::index,
        "MATCH" => lookups::match_,
        "XLOOKUP" => lookups::xlookup,
        "CHOOSE" => lookups::choose,
        "ROW" => lookups::row,
        "COLUMN" => lookups::column,
        "ROWS" => lookups::rows,
        "COLUMNS" => lookups::columns,
        "ISBLANK" => info::isblank,
        "ISNUMBER" => info::isnumber,
        "ISTEXT" => info::istext,
        "ISLOGICAL" => info::islogical,
        "ISERROR" => info::iserror,
        "ISERR" => info::iserr,
        "ISNA" => info::isna,
        "NA" => info::na,
        "N" => info::n,
        _ => return None,
    })
}

/// collect numbers for aggregation: referenced cells contribute only numeric
/// values, literal/computed arguments coerce, errors propagate.
pub(crate) fn collect_numbers(
    args: &[Expr],
    ctx: &EvalContext<'_>,
) -> Result<Vec<f64>, ErrorValue> {
    let mut nums = Vec::new();
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                for value in area.values(ctx)? {
                    push_reference_number(&mut nums, value)?;
                }
            }
            None => match evaluate(arg, ctx) {
                CellValue::Number { value } => nums.push(value),
                CellValue::Bool { value } => nums.push(if value { 1.0 } else { 0.0 }),
                CellValue::Empty => {}
                CellValue::Text { value } => match parse_num(&value) {
                    Some(n) => nums.push(n),
                    None => return Err(ErrorValue::Value),
                },
                CellValue::Error { value } => return Err(value),
            },
        }
    }
    Ok(nums)
}

/// a referenced cell contributes to aggregation only when numeric; errors
/// propagate, text/bool/blank are silently skipped.
fn push_reference_number(nums: &mut Vec<f64>, v: CellValue) -> Result<(), ErrorValue> {
    match v {
        CellValue::Number { value } => nums.push(value),
        CellValue::Error { value } => return Err(value),
        _ => {}
    }
    Ok(())
}

/// evaluate one argument and coerce it to a number, propagating errors.
pub(crate) fn nth_number(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    i: usize,
) -> Result<f64, ErrorValue> {
    to_number(&evaluate(&args[i], ctx))
}

/// evaluate one argument, coerce to a number, truncate toward zero.
pub(crate) fn nth_int(args: &[Expr], ctx: &EvalContext<'_>, i: usize) -> Result<i64, ErrorValue> {
    Ok(nth_number(args, ctx, i)?.trunc() as i64)
}

/// finalize a computed float: non-finite results become `#NUM!`.
pub(crate) fn finite(x: f64) -> CellValue {
    if x.is_finite() {
        num(x)
    } else {
        err(ErrorValue::Num)
    }
}

#[cfg(test)]
mod tests {
    use super::Params;

    #[test]
    fn a_criterion_follows_each_range_it_applies_to() {
        let countifs = Params::CriteriaFrom(1);
        assert!([1, 3, 5].iter().all(|i| countifs.takes_a_value(*i)));
        assert!([0, 2, 4].iter().all(|i| !countifs.takes_a_value(*i)));
        let sumifs = Params::CriteriaFrom(2);
        assert!([2, 4].iter().all(|i| sumifs.takes_a_value(*i)));
        assert!([0, 1, 3].iter().all(|i| !sumifs.takes_a_value(*i)));
    }

    #[test]
    fn a_function_nobody_described_reads_ranges_as_it_always_did() {
        assert_eq!(super::params("SUM"), Params::Ranges);
        assert_eq!(super::params("NOT_A_FUNCTION"), Params::Ranges);
        assert_eq!(super::params("len"), Params::Values);
    }
}
