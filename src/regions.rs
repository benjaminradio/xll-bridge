//! Formula regions: a block of cells that must all carry the same formula (the same R1C1 formula,
//! which is exactly what "filled right / filled down" produces).
//!
//! A region is a hidden workbook name (`_xs_<label>`) pointing at the block, so Excel itself
//! keeps it current when rows or columns are inserted or deleted. The block's top-left cell is
//! the **anchor**: its formula is the definition. The pieces here are pure so they are tested
//! without Excel; the object-model work is in `xll/ribbon/regions.rs`.
//!
//! The workbook stays an ordinary workbook: the cells hold plain formulas, so nothing has to be
//! stripped before it is shared with someone who has no add-in.

use crate::value::Value;

/// Prefix of the hidden names that mark regions.
pub const NAME_PREFIX: &str = "_xs_";

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RegionInfo {
    /// The label (the name without its prefix).
    pub name: String,
    /// `Sheet!$E$5:$KJ$5`
    pub address: String,
    pub rows: u32,
    pub cols: u32,
    /// How the region follows a workbook-wide roll-forward.
    pub rule: Rule,
}

/// How a region moves when the whole workbook rolls forward by one period.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum RollMode {
    /// Stays where it is.
    #[default]
    None,
    /// The start moves (a forecast: its first period became an actual).
    Start,
    /// The end moves (actuals: one more period).
    End,
    /// Both move (a rolling forecast).
    Slide,
}

impl RollMode {
    pub fn word(self) -> &'static str {
        match self {
            RollMode::None => "none",
            RollMode::Start => "start",
            RollMode::End => "end",
            RollMode::Slide => "slide",
        }
    }

    pub fn from_word(w: &str) -> Option<RollMode> {
        match w.trim().to_ascii_lowercase().as_str() {
            "none" | "fixed" | "" => Some(RollMode::None),
            "start" | "forecast" => Some(RollMode::Start),
            "end" | "actuals" | "extend" => Some(RollMode::End),
            "slide" | "both" | "rolling" => Some(RollMode::Slide),
            _ => None,
        }
    }
}

/// What happens to cells that leave a region when it rolls.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum OldCells {
    /// They become values (the usual roll-forward).
    #[default]
    Values,
    /// They keep their formulas.
    Keep,
    /// They are cleared.
    Clear,
}

impl OldCells {
    pub fn word(self) -> &'static str {
        match self {
            OldCells::Values => "values",
            OldCells::Keep => "keep",
            OldCells::Clear => "clear",
        }
    }

    pub fn from_word(w: &str) -> Option<OldCells> {
        match w.trim().to_ascii_lowercase().as_str() {
            "values" | "value" | "freeze" => Some(OldCells::Values),
            "keep" | "formulas" => Some(OldCells::Keep),
            "clear" | "cleared" | "blank" => Some(OldCells::Clear),
            _ => None,
        }
    }
}

/// A region's roll rule; stored in the comment of its hidden name (`roll=slide;old=clear`).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Rule {
    pub mode: RollMode,
    /// What becomes of cells that leave the region.
    pub old: OldCells,
}

impl Rule {
    pub fn encode(self) -> String {
        format!("roll={};old={}", self.mode.word(), self.old.word())
    }

    /// Reads a name's comment; anything unreadable means the default (no rule). The earlier
    /// `freeze=0/1` spelling still reads.
    pub fn parse(comment: &str) -> Rule {
        let mut rule = Rule::default();
        for part in comment.split(';') {
            match part.trim().split_once('=') {
                Some(("roll", w)) => rule.mode = RollMode::from_word(w).unwrap_or_default(),
                Some(("old", w)) => rule.old = OldCells::from_word(w).unwrap_or_default(),
                Some(("freeze", v)) => {
                    rule.old = if v.trim() == "0" {
                        OldCells::Keep
                    } else {
                        OldCells::Values
                    }
                }
                _ => {}
            }
        }
        rule
    }

    /// `(start, end)` for rolling `n` periods forward (negative = back).
    pub fn steps(self, n: i32) -> (i32, i32) {
        match self.mode {
            RollMode::None => (0, 0),
            RollMode::Start => (n, 0),
            RollMode::End => (0, n),
            RollMode::Slide => (n, n),
        }
    }

    pub fn describe(self) -> String {
        match self.mode {
            RollMode::None => "does not roll".into(),
            RollMode::Start => format!("moves its start{}", self.old_note()),
            RollMode::End => "extends its end".into(),
            RollMode::Slide => format!("slides its window{}", self.old_note()),
        }
    }

    fn old_note(self) -> &'static str {
        match self.old {
            OldCells::Values => " (old cells become values)",
            OldCells::Keep => " (old cells keep formulas)",
            OldCells::Clear => " (old cells are cleared)",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DevKind {
    /// Nothing in the cell.
    Blank,
    /// A typed value instead of a formula.
    Constant,
    /// A different formula.
    Formula,
}

impl DevKind {
    pub fn word(self) -> &'static str {
        match self {
            DevKind::Blank => "blank",
            DevKind::Constant => "constant",
            DevKind::Formula => "formula",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Deviation {
    pub address: String,
    pub kind: DevKind,
    pub text: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CheckReport {
    pub region: RegionInfo,
    /// The anchor's R1C1 formula, which every cell must carry.
    pub anchor: String,
    pub deviations: Vec<Deviation>,
}

impl CheckReport {
    /// Deviations that hold something (blank cells are simply filled in).
    pub fn non_blank(&self) -> usize {
        self.deviations
            .iter()
            .filter(|d| d.kind != DevKind::Blank)
            .count()
    }
}

/// `forecast dates` -> `_xs_forecast_dates`; the label is cleaned to letters, digits, `_` and `.`.
pub fn name_for(label: &str) -> Result<String, String> {
    let clean: String = label
        .trim()
        .chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || c == '_' || c == '.' {
                c
            } else {
                '_'
            }
        })
        .collect();
    if clean.is_empty() || clean.len() > 200 {
        return Err("a region needs a name of 1 to 200 characters".into());
    }
    Ok(format!("{NAME_PREFIX}{clean}"))
}

/// The label of a hidden region name, if it is one.
pub fn label_of(name: &str) -> Option<&str> {
    let n = name.rsplit('!').next().unwrap_or(name);
    n.strip_prefix(NAME_PREFIX)
}

/// What a cell holds, for comparison: its `FormulaR1C1` as a value.
pub fn classify(v: &Value) -> (DevKind, String) {
    match v {
        Value::Empty => (DevKind::Blank, String::new()),
        Value::Str(s) if s.is_empty() => (DevKind::Blank, String::new()),
        Value::Str(s) if s.starts_with('=') => (DevKind::Formula, s.clone()),
        Value::Str(s) => (DevKind::Constant, s.clone()),
        Value::Num(n) => (DevKind::Constant, format!("{n}")),
        Value::Bool(b) => (DevKind::Constant, b.to_string().to_uppercase()),
        Value::Err(e) => (DevKind::Constant, e.excel_text().to_owned()),
        Value::Array(_) => (DevKind::Constant, String::new()),
    }
}

/// Every cell of the row-major block that does not carry `anchor`. `address_of(row, col)` names a
/// cell (0-based within the block).
pub fn compare(
    anchor: &str,
    cells: &[Value],
    cols: usize,
    address_of: impl Fn(usize, usize) -> String,
) -> Vec<Deviation> {
    let mut out = Vec::new();
    for (i, v) in cells.iter().enumerate() {
        let (kind, text) = classify(v);
        if kind == DevKind::Formula && text == anchor {
            continue;
        }
        out.push(Deviation {
            address: address_of(i / cols.max(1), i % cols.max(1)),
            kind,
            text,
        });
    }
    out
}

// ------------------------------------------------------------------ rolling

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Axis {
    Columns,
    Rows,
}

/// Along which axis a block rolls: its long side (a row of periods rolls along its columns).
pub fn default_axis(rows: i32, cols: i32) -> Axis {
    if rows == 1 || cols >= rows {
        Axis::Columns
    } else {
        Axis::Rows
    }
}

/// A block of cells given relative to the region's original top-left cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Rect {
    pub row: i32,
    pub col: i32,
    pub rows: i32,
    pub cols: i32,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Roll {
    /// The region after the change.
    pub new: Rect,
    /// Cells that stop belonging to it.
    pub leaving: Vec<Rect>,
    /// Cells that start belonging to it.
    pub entering: Vec<Rect>,
}

/// Moves the start of the region by `start` cells and its end by `end` cells along `axis`
/// (positive = forward). Start forward + end 0 = move the start ("forecast: a period became an
/// actual"), start 0 + end forward = extend the end ("actuals: one more period"), both = slide
/// the window ("rolling forecast").
pub fn roll(rows: i32, cols: i32, axis: Axis, start: i32, end: i32) -> Result<Roll, String> {
    let len = if axis == Axis::Columns { cols } else { rows };
    let (ns, ne) = (start, len + end);
    if ne <= ns {
        return Err("the region would be left with no cells".into());
    }
    let along = |a: i32, b: i32| -> Rect {
        match axis {
            Axis::Columns => Rect {
                row: 0,
                col: a,
                rows,
                cols: b - a,
            },
            Axis::Rows => Rect {
                row: a,
                col: 0,
                rows: b - a,
                cols,
            },
        }
    };
    let mut leaving = Vec::new();
    let mut entering = Vec::new();
    let a_end = ns.clamp(0, len);
    if a_end > 0 {
        leaving.push(along(0, a_end));
    }
    let b_start = ne.clamp(0, len);
    if b_start < len {
        leaving.push(along(b_start, len));
    }
    if ns < ne.min(0) {
        entering.push(along(ns, ne.min(0)));
    }
    if ns.max(len) < ne {
        entering.push(along(ns.max(len), ne));
    }
    Ok(Roll {
        new: along(ns, ne),
        leaving,
        entering,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn names_are_cleaned_and_recognised() {
        assert_eq!(name_for("Forecast dates").unwrap(), "_xs_Forecast_dates");
        assert_eq!(label_of("_xs_Forecast_dates"), Some("Forecast_dates"));
        assert_eq!(label_of("Sheet1!_xs_x"), Some("x"));
        assert_eq!(label_of("Sales"), None);
        assert!(name_for("  ").is_err());
    }

    fn a1(r: usize, c: usize) -> String {
        format!("R{}C{}", r + 1, c + 1)
    }

    #[test]
    fn deviations_are_blank_constant_or_different_formula() {
        let f = |s: &str| Value::Str(s.into());
        let cells = vec![
            f("=EOMONTH(RC[-1],1)"),
            f("=EOMONTH(RC[-1],1)"),
            Value::Num(42.0),
            Value::Empty,
            f("=EOMONTH(RC[-1],2)"),
            f("'note"),
        ];
        let d = compare("=EOMONTH(RC[-1],1)", &cells, 3, a1);
        let seen: Vec<_> = d.iter().map(|x| (x.address.as_str(), x.kind)).collect();
        assert_eq!(
            seen,
            [
                ("R1C3", DevKind::Constant),
                ("R2C1", DevKind::Blank),
                ("R2C2", DevKind::Formula),
                ("R2C3", DevKind::Constant)
            ]
        );
        assert_eq!(d[0].text, "42");
        let report = CheckReport {
            region: RegionInfo {
                name: "x".into(),
                address: "S!A1".into(),
                rows: 2,
                cols: 3,
                rule: Rule::default(),
            },
            anchor: String::new(),
            deviations: d,
        };
        assert_eq!(report.non_blank(), 3);
        assert!(compare("=A1", &[Value::Str("=A1".into())], 1, a1).is_empty());
    }

    fn rect(row: i32, col: i32, rows: i32, cols: i32) -> Rect {
        Rect {
            row,
            col,
            rows,
            cols,
        }
    }

    #[test]
    fn moving_the_start_extending_the_end_and_sliding() {
        // A row of 10 periods.
        let (r, c) = (1, 10);
        // Forecast: the first period became an actual.
        let m = roll(r, c, Axis::Columns, 1, 0).unwrap();
        assert_eq!(m.new, rect(0, 1, 1, 9));
        assert_eq!(m.leaving, [rect(0, 0, 1, 1)]);
        assert!(m.entering.is_empty());
        // Actuals: one more period.
        let e = roll(r, c, Axis::Columns, 0, 1).unwrap();
        assert_eq!(e.new, rect(0, 0, 1, 11));
        assert_eq!(e.entering, [rect(0, 10, 1, 1)]);
        assert!(e.leaving.is_empty());
        // Rolling forecast: the window slides by one.
        let s = roll(r, c, Axis::Columns, 1, 1).unwrap();
        assert_eq!(s.new, rect(0, 1, 1, 10));
        assert_eq!(
            (s.leaving, s.entering),
            (vec![rect(0, 0, 1, 1)], vec![rect(0, 10, 1, 1)])
        );
        // Backwards, and by several; a block of rows along rows.
        let back = roll(r, c, Axis::Columns, -2, -1).unwrap();
        assert_eq!(back.new, rect(0, -2, 1, 11));
        assert_eq!(back.entering, [rect(0, -2, 1, 2)]);
        assert_eq!(back.leaving, [rect(0, 9, 1, 1)]);
        let down = roll(5, 3, Axis::Rows, 2, 2).unwrap();
        assert_eq!(down.new, rect(2, 0, 5, 3));
        assert_eq!(down.entering, [rect(5, 0, 2, 3)]);
        assert_eq!(down.leaving, [rect(0, 0, 2, 3)]);
        // A slide further than the window: nothing of the old block stays.
        let far = roll(r, 3, Axis::Columns, 5, 5).unwrap();
        assert_eq!(far.new, rect(0, 5, 1, 3));
        assert_eq!(far.leaving, [rect(0, 0, 1, 3)]);
        assert_eq!(far.entering, [rect(0, 5, 1, 3)]);
        // Never an empty region.
        assert!(roll(r, c, Axis::Columns, 10, 0).is_err());
        assert!(roll(r, c, Axis::Columns, 0, -10).is_err());
        assert!(roll(r, c, Axis::Columns, 3, -7).is_err());
        assert!(roll(r, c, Axis::Columns, 3, -6).is_ok());
    }

    #[test]
    fn rules_round_trip_and_give_steps() {
        let r = Rule {
            mode: RollMode::Slide,
            old: OldCells::Clear,
        };
        assert_eq!(r.encode(), "roll=slide;old=clear");
        assert_eq!(Rule::parse(&r.encode()), r);
        assert_eq!(Rule::parse("roll=start;freeze=0").old, OldCells::Keep);
        assert_eq!(Rule::parse("roll=start;freeze=1").old, OldCells::Values);
        assert_eq!(Rule::parse(""), Rule::default());
        assert_eq!(Rule::parse("junk;roll=bogus"), Rule::default());
        assert_eq!(Rule::parse("roll=forecast").mode, RollMode::Start);
        assert_eq!(Rule::parse("roll=end").old, OldCells::Values);
        assert_eq!(r.steps(1), (1, 1));
        assert_eq!(r.steps(-1), (-1, -1));
        let s = |m| Rule {
            mode: m,
            old: OldCells::Values,
        };
        assert_eq!(s(RollMode::Start).steps(1), (1, 0));
        assert_eq!(s(RollMode::End).steps(1), (0, 1));
        assert_eq!(s(RollMode::None).steps(1), (0, 0));
        assert_eq!(RollMode::from_word("Rolling"), Some(RollMode::Slide));
        assert_eq!(RollMode::from_word("nonsense"), None);
    }

    #[test]
    fn the_long_side_is_the_default_axis() {
        assert_eq!(default_axis(1, 359), Axis::Columns);
        assert_eq!(default_axis(40, 1), Axis::Rows);
        assert_eq!(default_axis(4, 12), Axis::Columns);
        assert_eq!(default_axis(12, 4), Axis::Rows);
    }
}
