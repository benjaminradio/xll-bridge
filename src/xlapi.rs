//! Reading and writing worksheet cells on behalf of Lua.
//!
//! Lua code runs either on Excel's UI thread (`LUA.CALL`) or on background workers
//! (`LUA.ASYNC`, macros), but Excel's object model may only be used from the UI thread. So Lua
//! talks to an [`XlApi`]; on the UI thread a call is executed directly, on any other thread it is
//! queued to the UI thread through a [`Bridge`] and the caller waits for the answer.
//!
//! This module is independent of Excel: the executor that really talks to Excel is supplied by
//! the Windows layer, and tests use an in-memory grid.

use std::collections::VecDeque;
use std::path::PathBuf;
use std::sync::mpsc;
use std::sync::{Arc, Mutex, OnceLock};
use std::thread::ThreadId;
use std::time::Duration;

use crate::regions::{CheckReport, RegionInfo};
use crate::value::Value;

/// Largest range (in cells) read or written in one operation.
pub const MAX_CELLS: u64 = 5_000_000;
/// How long a worker waits for the UI thread to answer.
pub const CALL_TIMEOUT: Duration = Duration::from_secs(60);
/// A busy Excel (editing a cell, showing a dialog) is retried this many pumps before giving up.
const MAX_ATTEMPTS: u32 = 40;

#[derive(Clone, Debug)]
pub enum Op {
    /// Resolve an address: normalised form and size.
    Info {
        workbook: PathBuf,
        address: String,
    },
    Read {
        workbook: PathBuf,
        address: String,
        formulas: bool,
    },
    /// Writes `data` with its top-left cell at the range's top-left cell; an array sets exactly
    /// its own size, a scalar fills the whole range.
    Write {
        workbook: PathBuf,
        address: String,
        formulas: bool,
        data: Value,
    },
    Clear {
        workbook: PathBuf,
        address: String,
    },
    /// An Excel table (ListObject): header row and data, by table name.
    Table {
        workbook: PathBuf,
        name: String,
    },
    /// Turns a range into an Excel table (a single cell stands for its current region, like
    /// Ctrl+T), or resizes the table of that name on the sheet to it.
    AsTable {
        workbook: PathBuf,
        address: String,
        name: String,
    },
    /// Formula regions (see [`crate::regions`]).
    Region {
        workbook: PathBuf,
        op: RegionOp,
    },
    Sheets {
        workbook: PathBuf,
    },
    UsedRange {
        workbook: PathBuf,
        sheet: String,
    },
    Selection {
        workbook: PathBuf,
    },
}

/// Operations on formula regions. A region is named by its label.
#[derive(Clone, Debug)]
pub enum RegionOp {
    /// Make the range a region (its top-left cell is the anchor) and fill it with the anchor's
    /// formula. Cells that hold something else stop this unless `overwrite`.
    Define {
        address: String,
        label: Option<String>,
        overwrite: bool,
    },
    List,
    /// The region containing a cell, if any.
    At {
        address: String,
    },
    Check {
        name: String,
    },
    /// Fill the region from its anchor again.
    Apply {
        name: String,
        overwrite: bool,
    },
    /// Move the start / end of the region along its long side; new cells get the anchor's formula
    /// and the cells that leave it become values, stay as they are or are cleared (`old`).
    Roll {
        name: String,
        start: i32,
        end: i32,
        old: crate::regions::OldCells,
        overwrite: bool,
    },
    /// Forget the region; its cells stay as they are.
    Remove {
        name: String,
    },
    /// Set how the region follows a workbook-wide roll-forward.
    SetRule {
        name: String,
        rule: crate::regions::Rule,
    },
}

impl Op {
    pub fn is_write(&self) -> bool {
        match self {
            Op::Write { .. } | Op::Clear { .. } | Op::AsTable { .. } => true,
            Op::Region { op, .. } => !matches!(
                op,
                RegionOp::List | RegionOp::At { .. } | RegionOp::Check { .. }
            ),
            _ => false,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RangeInfo {
    pub address: String,
    pub rows: u32,
    pub cols: u32,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Reply {
    Regions(Vec<RegionInfo>),
    Report(CheckReport),
    Info(RangeInfo),
    Data(Value),
    Names(Vec<String>),
    Done,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct XlError {
    pub message: String,
    /// Excel was busy; the operation may succeed if tried again shortly.
    pub retryable: bool,
}

impl XlError {
    pub fn new(message: impl Into<String>) -> Self {
        XlError {
            message: message.into(),
            retryable: false,
        }
    }
    pub fn busy(message: impl Into<String>) -> Self {
        XlError {
            message: message.into(),
            retryable: true,
        }
    }
}

impl std::fmt::Display for XlError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.message)
    }
}

/// The message for cells that stop a region from being filled: what they hold and where.
pub fn describe_deviations(report: &CheckReport) -> String {
    let held: Vec<String> = report
        .deviations
        .iter()
        .filter(|d| d.kind != crate::regions::DevKind::Blank)
        .map(|d| {
            format!(
                "{} ({}{})",
                d.address,
                d.kind.word(),
                if d.text.is_empty() {
                    String::new()
                } else {
                    format!(": {}", d.text)
                }
            )
        })
        .collect();
    let shown = held.iter().take(8).cloned().collect::<Vec<_>>().join(", ");
    let more = held.len().saturating_sub(8);
    format!(
        "{} cell(s) of region '{}' hold something other than its formula {}: {shown}{}; pass overwrite = true to replace them",
        held.len(),
        report.region.name,
        report.anchor,
        if more > 0 {
            format!(" and {more} more")
        } else {
            String::new()
        }
    )
}

pub trait XlApi: Send + Sync {
    fn call(&self, op: Op) -> Result<Reply, XlError>;
}

/// Refuses to change cells: what `LUA.CALL` gets, because Excel does not allow a worksheet
/// function to modify the sheet it is calculating.
pub struct ReadOnly(pub Arc<dyn XlApi>);

impl XlApi for ReadOnly {
    fn call(&self, op: Op) -> Result<Reply, XlError> {
        if op.is_write() {
            return Err(XlError::new(
                "cells cannot be changed from LUA.CALL (Excel forbids it); use LUA.ASYNC or a macro (xl.macro)",
            ));
        }
        self.0.call(op)
    }
}

// ------------------------------------------------------------------ the thread bridge

pub type Executor = dyn Fn(Op) -> Result<Reply, XlError> + Send + Sync;

struct Request {
    op: Op,
    reply: mpsc::Sender<Result<Reply, XlError>>,
    attempts: u32,
}

pub struct Bridge {
    exec: Arc<Executor>,
    ui_thread: ThreadId,
    queue: Mutex<VecDeque<Request>>,
    wake: Box<dyn Fn() + Send + Sync>,
    timeout: Duration,
}

impl Bridge {
    /// Must be created on the UI thread: that is where requests are executed.
    pub fn new(exec: Arc<Executor>, wake: Box<dyn Fn() + Send + Sync>, timeout: Duration) -> Self {
        Bridge {
            exec,
            ui_thread: std::thread::current().id(),
            queue: Mutex::new(VecDeque::new()),
            wake,
            timeout,
        }
    }

    /// Runs the queued requests. Call on the UI thread (from a posted message or a timer).
    /// A request that hit a busy Excel stays queued for the next pump. Returns how many finished.
    pub fn pump(&self) -> usize {
        debug_assert_eq!(
            std::thread::current().id(),
            self.ui_thread,
            "pump must run on the UI thread"
        );
        let n = self.queue.lock().unwrap_or_else(|e| e.into_inner()).len();
        let mut finished = 0;
        for _ in 0..n {
            let Some(mut req) = self
                .queue
                .lock()
                .unwrap_or_else(|e| e.into_inner())
                .pop_front()
            else {
                break;
            };
            let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                (self.exec)(req.op.clone())
            }))
            .unwrap_or_else(|_| Err(XlError::new("internal error while talking to Excel")));
            match outcome {
                Err(e) if e.retryable && req.attempts + 1 < MAX_ATTEMPTS => {
                    req.attempts += 1;
                    self.queue
                        .lock()
                        .unwrap_or_else(|e| e.into_inner())
                        .push_back(req);
                }
                other => {
                    let _ = req.reply.send(other); // the waiter may have timed out and gone
                    finished += 1;
                }
            }
        }
        finished
    }

    pub fn queued(&self) -> usize {
        self.queue.lock().unwrap_or_else(|e| e.into_inner()).len()
    }
}

impl XlApi for Bridge {
    fn call(&self, op: Op) -> Result<Reply, XlError> {
        if std::thread::current().id() == self.ui_thread {
            return (self.exec)(op);
        }
        let (tx, rx) = mpsc::channel();
        self.queue
            .lock()
            .unwrap_or_else(|e| e.into_inner())
            .push_back(Request {
                op,
                reply: tx,
                attempts: 0,
            });
        (self.wake)();
        rx.recv_timeout(self.timeout).unwrap_or_else(|_| {
            Err(XlError::new(
                "Excel did not respond (it may be busy with a dialog or a long calculation)",
            ))
        })
    }
}

static GLOBAL: OnceLock<Arc<Bridge>> = OnceLock::new();

/// Installs the process-wide bridge. Only the first call has any effect.
pub fn init(bridge: Bridge) {
    let _ = GLOBAL.set(Arc::new(bridge));
}

/// The bridge to Excel, if the add-in has connected (it has not in tests or the packer).
pub fn global() -> Option<Arc<dyn XlApi>> {
    GLOBAL.get().map(|b| Arc::clone(b) as Arc<dyn XlApi>)
}

/// Runs queued requests of the global bridge (UI thread).
pub fn pump_global() -> usize {
    GLOBAL.get().map_or(0, |b| b.pump())
}

// ------------------------------------------------------------------ addresses

/// `[Book.xlsx]Sheet 1!A1:B2` and its quoted forms, taken apart.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Addr {
    pub book: Option<String>,
    pub sheet: Option<String>,
    /// The cell reference or defined name after the `!` (or the whole string when unqualified).
    pub cells: String,
}

pub fn split_address(input: &str) -> Result<Addr, String> {
    let s = input.trim();
    if s.is_empty() {
        return Err("empty address".into());
    }
    let (qualifier, cells) = if let Some(rest) = s.strip_prefix('\'') {
        // 'Sheet name'!A1  or  '[Book.xlsx]Sheet name'!A1  ('' is an escaped quote)
        let mut name = String::new();
        let mut chars = rest.char_indices().peekable();
        let mut end = None;
        while let Some((i, c)) = chars.next() {
            if c == '\'' {
                if matches!(chars.peek(), Some((_, '\''))) {
                    name.push('\'');
                    chars.next();
                } else {
                    end = Some(i + 1);
                    break;
                }
            } else {
                name.push(c);
            }
        }
        let end = end.ok_or_else(|| format!("unterminated quote in '{input}'"))?;
        let after = &rest[end..];
        let cells = after
            .strip_prefix('!')
            .ok_or_else(|| format!("expected '!' after the sheet name in '{input}'"))?;
        (Some(name), cells.to_owned())
    } else if let Some(bang) = bang_outside_brackets(s) {
        (Some(s[..bang].to_owned()), s[bang + 1..].to_owned())
    } else {
        (None, s.to_owned())
    };
    let cells = cells.trim().to_owned();
    if cells.is_empty() {
        return Err(format!("no cells after the sheet name in '{input}'"));
    }
    let (book, sheet) = match qualifier {
        None => (None, None),
        Some(q) => match q.strip_prefix('[').and_then(|r| r.split_once(']')) {
            Some((book, sheet)) => (
                Some(book.to_owned()),
                (!sheet.is_empty()).then(|| sheet.to_owned()),
            ),
            None => (None, Some(q)),
        },
    };
    Ok(Addr { book, sheet, cells })
}

fn bang_outside_brackets(s: &str) -> Option<usize> {
    let mut depth = 0;
    for (i, c) in s.char_indices() {
        match c {
            '[' => depth += 1,
            ']' => depth -= 1,
            '!' if depth == 0 => return Some(i),
            _ => {}
        }
    }
    None
}

/// `Sheet 1` -> `'Sheet 1'`, leaving simple names bare.
pub fn quote_sheet(name: &str) -> String {
    let simple = !name.is_empty()
        && name.chars().all(|c| c.is_alphanumeric() || c == '_')
        && !name.chars().next().is_some_and(|c| c.is_ascii_digit());
    if simple {
        name.to_owned()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

#[cfg(test)]
pub(crate) mod fake {
    //! An in-memory worksheet standing in for Excel.
    use super::*;
    use std::collections::{BTreeMap, HashMap};

    #[derive(Default)]
    pub struct Sheet {
        pub values: HashMap<(u32, u32), Value>,
        pub formulas: HashMap<(u32, u32), String>,
    }

    pub struct MemGrid {
        pub sheets: Mutex<BTreeMap<String, Sheet>>,
        pub active: String,
        pub selection: Mutex<Option<String>>,
        pub log: Mutex<Vec<String>>,
        /// Excel table name -> address of its range (header included).
        pub tables: Mutex<HashMap<String, String>>,
        /// Formula regions: label -> (sheet, (row1, col1, row2, col2)), 0-based.
        pub regions: Mutex<BTreeMap<String, (String, Bounds)>>,
        pub rules: Mutex<BTreeMap<String, crate::regions::Rule>>,
    }

    /// (row1, col1, row2, col2), 0-based.
    pub type Bounds = (u32, u32, u32, u32);

    pub fn col_number(letters: &str) -> Option<u32> {
        if letters.is_empty() {
            return None;
        }
        letters.chars().try_fold(0u32, |n, c| {
            c.is_ascii_alphabetic()
                .then(|| n * 26 + (c.to_ascii_uppercase() as u32 - 64))
        })
    }

    fn cell(s: &str) -> Option<(u32, u32)> {
        let s = s.replace('$', "");
        let split = s.find(|c: char| c.is_ascii_digit())?;
        let col = col_number(&s[..split])?;
        let row: u32 = s[split..].parse().ok()?;
        Some((row - 1, col - 1))
    }

    /// `A1`, `A1:C3` -> (row1, col1, row2, col2), 0-based.
    pub fn parse_range(s: &str) -> Option<Bounds> {
        let (a, b) = s.split_once(':').unwrap_or((s, s));
        let (r1, c1) = cell(a)?;
        let (r2, c2) = cell(b)?;
        Some((r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)))
    }

    pub fn a1(row: u32, col: u32) -> String {
        let mut c = col + 1;
        let mut l = Vec::new();
        while c > 0 {
            l.push(b'A' + ((c - 1) % 26) as u8);
            c = (c - 1) / 26;
        }
        l.reverse();
        format!("{}{}", String::from_utf8_lossy(&l), row + 1)
    }

    impl MemGrid {
        pub fn new() -> Arc<Self> {
            let mut sheets = BTreeMap::new();
            sheets.insert("Data".to_owned(), Sheet::default());
            sheets.insert("Other".to_owned(), Sheet::default());
            Arc::new(MemGrid {
                sheets: Mutex::new(sheets),
                active: "Data".into(),
                selection: Mutex::new(None),
                log: Mutex::default(),
                tables: Mutex::default(),
                regions: Mutex::default(),
                rules: Mutex::default(),
            })
        }

        fn region_info(&self, label: &str) -> Result<crate::regions::RegionInfo, XlError> {
            let regions = self.regions.lock().unwrap();
            let (sheet, (r1, c1, r2, c2)) = regions
                .get(label)
                .ok_or_else(|| XlError::new(format!("no formula region named '{label}'")))?;
            Ok(crate::regions::RegionInfo {
                name: label.to_owned(),
                address: format!("{}!{}:{}", quote_sheet(sheet), a1(*r1, *c1), a1(*r2, *c2)),
                rows: r2 - r1 + 1,
                cols: c2 - c1 + 1,
                rule: self
                    .rules
                    .lock()
                    .unwrap()
                    .get(label)
                    .copied()
                    .unwrap_or_default(),
            })
        }

        fn cell(&self, sheet: &str, r: u32, c: u32) -> Value {
            let sheets = self.sheets.lock().unwrap();
            let s = &sheets[sheet];
            match s.formulas.get(&(r, c)) {
                Some(f) => Value::Str(f.clone()),
                None => s.values.get(&(r, c)).cloned().unwrap_or(Value::Empty),
            }
        }

        fn check_region(&self, label: &str) -> Result<crate::regions::CheckReport, XlError> {
            let region = self.region_info(label)?;
            let (sheet, (r1, c1, r2, c2)) = self.regions.lock().unwrap()[label].clone();
            let anchor = match self.cell(&sheet, r1, c1) {
                Value::Str(f) if f.starts_with('=') => f,
                _ => return Err(XlError::new("the anchor (top-left cell) holds no formula")),
            };
            let mut cells = Vec::new();
            for r in r1..=r2 {
                for c in c1..=c2 {
                    cells.push(self.cell(&sheet, r, c));
                }
            }
            let deviations =
                crate::regions::compare(&anchor, &cells, (c2 - c1 + 1) as usize, |i, j| {
                    a1(r1 + i as u32, c1 + j as u32)
                });
            Ok(crate::regions::CheckReport {
                region,
                anchor,
                deviations,
            })
        }

        fn apply_region(&self, label: &str, overwrite: bool) -> Result<(), XlError> {
            let report = self.check_region(label)?;
            if report.non_blank() > 0 && !overwrite {
                return Err(XlError::new(describe_deviations(&report)));
            }
            let (sheet, (r1, c1, r2, c2)) = self.regions.lock().unwrap()[label].clone();
            let mut sheets = self.sheets.lock().unwrap();
            let s = sheets.get_mut(&sheet).unwrap();
            for r in r1..=r2 {
                for c in c1..=c2 {
                    s.values.remove(&(r, c));
                    s.formulas.insert((r, c), report.anchor.clone());
                }
            }
            Ok(())
        }

        fn region(&self, op: RegionOp) -> Result<Reply, XlError> {
            use crate::regions::{Axis, default_axis, label_of, name_for, roll};
            match op {
                RegionOp::List => {
                    let labels: Vec<String> =
                        self.regions.lock().unwrap().keys().cloned().collect();
                    labels
                        .iter()
                        .map(|l| self.region_info(l))
                        .collect::<Result<Vec<_>, _>>()
                        .map(Reply::Regions)
                }
                RegionOp::At { address } => {
                    let (sheet, (r, c, _, _)) = self.resolve(&address)?;
                    let hit = self
                        .regions
                        .lock()
                        .unwrap()
                        .iter()
                        .find(|(_, (s, (r1, c1, r2, c2)))| {
                            *s == sheet && r >= *r1 && r <= *r2 && c >= *c1 && c <= *c2
                        })
                        .map(|(l, _)| l.clone());
                    match hit {
                        Some(l) => Ok(Reply::Regions(vec![self.region_info(&l)?])),
                        None => Ok(Reply::Regions(Vec::new())),
                    }
                }
                RegionOp::Check { name } => self.check_region(&name).map(Reply::Report),
                RegionOp::SetRule { name, rule } => {
                    self.region_info(&name)?;
                    self.rules.lock().unwrap().insert(name.clone(), rule);
                    Ok(Reply::Regions(vec![self.region_info(&name)?]))
                }
                RegionOp::Remove { name } => {
                    self.rules.lock().unwrap().remove(&name);
                    self.regions
                        .lock()
                        .unwrap()
                        .remove(&name)
                        .ok_or_else(|| XlError::new(format!("no formula region named '{name}'")))?;
                    Ok(Reply::Done)
                }
                RegionOp::Apply { name, overwrite } => {
                    self.apply_region(&name, overwrite)?;
                    Ok(Reply::Regions(vec![self.region_info(&name)?]))
                }
                RegionOp::Define {
                    address,
                    label,
                    overwrite,
                } => {
                    let (sheet, b @ (r1, c1, r2, c2)) = self.resolve(&address)?;
                    // A selection that overlaps one existing region redefines it (keeping its name and
                    // rule); one that overlaps several is refused.
                    let hits: Vec<(String, Bounds)> = self
                        .regions
                        .lock()
                        .unwrap()
                        .iter()
                        .filter(|(_, (s, (a, b2, c, d)))| {
                            *s == sheet && r1 <= *c && r2 >= *a && c1 <= *d && c2 >= *b2
                        })
                        .map(|(l, (_, old))| (l.clone(), *old))
                        .collect();
                    if hits.len() > 1 {
                        let names: Vec<&str> = hits.iter().map(|(l, _)| l.as_str()).collect();
                        return Err(XlError::new(format!(
                            "the range overlaps several formula regions ({}); remove them first",
                            names.join(", ")
                        )));
                    }
                    let same_anchor = hits.into_iter().next();
                    let old_anchor = same_anchor
                        .as_ref()
                        .and_then(|(l, _)| self.check_region(l).ok().map(|r| r.anchor));
                    let label = match &same_anchor {
                        Some((l, _)) => l.clone(),
                        None => label.unwrap_or_else(|| format!("{}_{}", sheet, a1(r1, c1))),
                    };
                    let label = label_of(&name_for(&label).map_err(XlError::new)?)
                        .unwrap_or_default()
                        .to_owned();
                    let overlaps =
                        self.regions
                            .lock()
                            .unwrap()
                            .iter()
                            .any(|(l, (s, (a, b2, c, d)))| {
                                *s == sheet
                                    && *l != label
                                    && r1 <= *c
                                    && r2 >= *a
                                    && c1 <= *d
                                    && c2 >= *b2
                            });
                    if overlaps {
                        return Err(XlError::new("the range overlaps another formula region"));
                    }
                    self.regions
                        .lock()
                        .unwrap()
                        .insert(label.clone(), (sheet.clone(), b));
                    // A new, blank anchor cell takes the old anchor's formula.
                    let mut wrote_anchor = false;
                    if let Some(f) = old_anchor {
                        let mut sheets = self.sheets.lock().unwrap();
                        let s = sheets.get_mut(&sheet).unwrap();
                        if !s.formulas.contains_key(&(r1, c1)) && !s.values.contains_key(&(r1, c1))
                        {
                            s.formulas.insert((r1, c1), f);
                            wrote_anchor = true;
                        }
                    }
                    if let Err(e) = self.apply_region(&label, overwrite) {
                        if wrote_anchor {
                            let mut sheets = self.sheets.lock().unwrap();
                            sheets.get_mut(&sheet).unwrap().formulas.remove(&(r1, c1));
                        }
                        let mut regions = self.regions.lock().unwrap();
                        match same_anchor {
                            Some((_, old)) => regions.insert(label.clone(), (sheet.clone(), old)),
                            None => regions.remove(&label),
                        };
                        return Err(e);
                    }
                    Ok(Reply::Regions(vec![self.region_info(&label)?]))
                }
                RegionOp::Roll {
                    name,
                    start,
                    end,
                    old,
                    overwrite,
                } => {
                    let (sheet, (r1, c1, r2, c2)) = self
                        .regions
                        .lock()
                        .unwrap()
                        .get(&name)
                        .cloned()
                        .ok_or_else(|| XlError::new(format!("no formula region named '{name}'")))?;
                    let (rows, cols) = ((r2 - r1 + 1) as i32, (c2 - c1 + 1) as i32);
                    let axis = default_axis(rows, cols);
                    let plan = roll(rows, cols, axis, start, end).map_err(XlError::new)?;
                    let _ = Axis::Rows;
                    let shifted = |rect: crate::regions::Rect| {
                        (
                            (r1 as i32 + rect.row) as u32,
                            (c1 as i32 + rect.col) as u32,
                            (r1 as i32 + rect.row + rect.rows - 1) as u32,
                            (c1 as i32 + rect.col + rect.cols - 1) as u32,
                        )
                    };
                    let anchor = self.check_region(&name)?.anchor;
                    if old != crate::regions::OldCells::Keep {
                        let mut sheets = self.sheets.lock().unwrap();
                        let s = sheets.get_mut(&sheet).unwrap();
                        for rect in &plan.leaving {
                            let (a, b, c, d) = shifted(*rect);
                            for r in a..=c {
                                for col in b..=d {
                                    // "Values": a formula becomes the (here: unknown) value.
                                    s.formulas.remove(&(r, col));
                                    if old == crate::regions::OldCells::Clear {
                                        s.values.remove(&(r, col));
                                    }
                                }
                            }
                        }
                    }
                    let nb = shifted(plan.new);
                    self.regions
                        .lock()
                        .unwrap()
                        .insert(name.clone(), (sheet.clone(), nb));
                    // The new cells get the anchor's formula; the new anchor is the first cell.
                    {
                        let mut sheets = self.sheets.lock().unwrap();
                        let s = sheets.get_mut(&sheet).unwrap();
                        for rect in &plan.entering {
                            let (a, b, c, d) = shifted(*rect);
                            for r in a..=c {
                                for col in b..=d {
                                    if !overwrite
                                        && (s.formulas.contains_key(&(r, col))
                                            || s.values.contains_key(&(r, col)))
                                    {
                                        continue;
                                    }
                                    s.values.remove(&(r, col));
                                    s.formulas.insert((r, col), anchor.clone());
                                }
                            }
                        }
                    }
                    Ok(Reply::Regions(vec![self.region_info(&name)?]))
                }
            }
        }

        pub fn put(&self, sheet: &str, address: &str, v: Value) {
            let (r, c, _, _) = parse_range(address).unwrap();
            self.sheets
                .lock()
                .unwrap()
                .get_mut(sheet)
                .unwrap()
                .values
                .insert((r, c), v);
        }

        pub fn get(&self, sheet: &str, address: &str) -> Value {
            let (r, c, _, _) = parse_range(address).unwrap();
            self.sheets.lock().unwrap()[sheet]
                .values
                .get(&(r, c))
                .cloned()
                .unwrap_or(Value::Empty)
        }

        fn resolve(&self, address: &str) -> Result<(String, Bounds), XlError> {
            let a = split_address(address).map_err(XlError::new)?;
            let sheet = a.sheet.unwrap_or_else(|| self.active.clone());
            if !self.sheets.lock().unwrap().contains_key(&sheet) {
                return Err(XlError::new(format!("no sheet named '{sheet}'")));
            }
            let r = parse_range(&a.cells)
                .ok_or_else(|| XlError::new(format!("'{}' is not a valid range", a.cells)))?;
            Ok((sheet, r))
        }
    }

    impl XlApi for MemGrid {
        fn call(&self, op: Op) -> Result<Reply, XlError> {
            self.log.lock().unwrap().push(format!("{op:?}"));
            match op {
                Op::Info { address, .. } => {
                    let (sheet, (r1, c1, r2, c2)) = self.resolve(&address)?;
                    Ok(Reply::Info(RangeInfo {
                        address: format!(
                            "{}!${}${}:${}${}",
                            quote_sheet(&sheet),
                            &a1(r1, c1)[..a1(r1, c1).find(|c: char| c.is_ascii_digit()).unwrap()],
                            r1 + 1,
                            &a1(r2, c2)[..a1(r2, c2).find(|c: char| c.is_ascii_digit()).unwrap()],
                            r2 + 1
                        ),
                        rows: r2 - r1 + 1,
                        cols: c2 - c1 + 1,
                    }))
                }
                Op::Read {
                    address, formulas, ..
                } => {
                    let (sheet, (r1, c1, r2, c2)) = self.resolve(&address)?;
                    let sheets = self.sheets.lock().unwrap();
                    let s = &sheets[&sheet];
                    let (rows, cols) = ((r2 - r1 + 1) as usize, (c2 - c1 + 1) as usize);
                    let mut data = Vec::with_capacity(rows * cols);
                    for r in r1..=r2 {
                        for c in c1..=c2 {
                            let v = if formulas {
                                s.formulas
                                    .get(&(r, c))
                                    .map(|f| Value::Str(f.clone()))
                                    .or_else(|| s.values.get(&(r, c)).cloned())
                            } else {
                                s.values.get(&(r, c)).cloned()
                            };
                            data.push(v.unwrap_or(Value::Empty));
                        }
                    }
                    Ok(Reply::Data(Value::Array(crate::value::Array::new(
                        rows, cols, data,
                    ))))
                }
                Op::Write {
                    address,
                    formulas,
                    data,
                    ..
                } => {
                    let (sheet, (r1, c1, r2, c2)) = self.resolve(&address)?;
                    let mut sheets = self.sheets.lock().unwrap();
                    let s = sheets.get_mut(&sheet).unwrap();
                    let mut put = |r: u32, c: u32, v: &Value| {
                        if formulas {
                            if let Value::Str(f) = v {
                                s.formulas.insert((r, c), f.clone());
                            }
                        } else {
                            s.formulas.remove(&(r, c));
                            s.values.insert((r, c), v.clone());
                        }
                    };
                    match &data {
                        Value::Array(a) => {
                            for i in 0..a.rows {
                                for j in 0..a.cols {
                                    put(r1 + i as u32, c1 + j as u32, &a.data[i * a.cols + j]);
                                }
                            }
                        }
                        scalar => {
                            for r in r1..=r2 {
                                for c in c1..=c2 {
                                    put(r, c, scalar);
                                }
                            }
                        }
                    }
                    Ok(Reply::Done)
                }
                Op::Clear { address, .. } => {
                    let (sheet, (r1, c1, r2, c2)) = self.resolve(&address)?;
                    let mut sheets = self.sheets.lock().unwrap();
                    let s = sheets.get_mut(&sheet).unwrap();
                    for r in r1..=r2 {
                        for c in c1..=c2 {
                            s.values.remove(&(r, c));
                            s.formulas.remove(&(r, c));
                        }
                    }
                    Ok(Reply::Done)
                }
                Op::AsTable { address, name, .. } => {
                    self.resolve(&address)?;
                    self.tables.lock().unwrap().insert(name, address);
                    Ok(Reply::Done)
                }
                Op::Region { op, .. } => self.region(op),
                Op::Table { name, .. } => {
                    let address = self
                        .tables
                        .lock()
                        .unwrap()
                        .get(&name)
                        .cloned()
                        .ok_or_else(|| XlError::new(format!("no table named '{name}'")))?;
                    self.call(Op::Read {
                        workbook: PathBuf::new(),
                        address,
                        formulas: false,
                    })
                }
                Op::Sheets { .. } => Ok(Reply::Names(
                    self.sheets.lock().unwrap().keys().cloned().collect(),
                )),
                Op::UsedRange { sheet, .. } => {
                    let sheets = self.sheets.lock().unwrap();
                    let s = sheets
                        .get(&sheet)
                        .ok_or_else(|| XlError::new(format!("no sheet named '{sheet}'")))?;
                    let keys: Vec<_> = s.values.keys().chain(s.formulas.keys()).collect();
                    let (r2, c2) = (
                        keys.iter().map(|k| k.0).max().unwrap_or(0),
                        keys.iter().map(|k| k.1).max().unwrap_or(0),
                    );
                    drop(sheets);
                    self.call(Op::Info {
                        workbook: PathBuf::new(),
                        address: format!("{}!A1:{}", quote_sheet(&sheet), a1(r2, c2)),
                    })
                }
                Op::Selection { .. } => match self.selection.lock().unwrap().clone() {
                    Some(a) => self.call(Op::Info {
                        workbook: PathBuf::new(),
                        address: a,
                    }),
                    None => Err(XlError::new("the selection is not a range of cells")),
                },
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn addr(book: Option<&str>, sheet: Option<&str>, cells: &str) -> Addr {
        Addr {
            book: book.map(str::to_owned),
            sheet: sheet.map(str::to_owned),
            cells: cells.to_owned(),
        }
    }

    #[test]
    fn addresses_in_every_form_excel_uses() {
        assert_eq!(split_address("A1:B2").unwrap(), addr(None, None, "A1:B2"));
        assert_eq!(
            split_address("Data!A1").unwrap(),
            addr(None, Some("Data"), "A1")
        );
        assert_eq!(
            split_address("'My Sheet'!A1:C3").unwrap(),
            addr(None, Some("My Sheet"), "A1:C3")
        );
        assert_eq!(
            split_address("'It''s'!B2").unwrap(),
            addr(None, Some("It's"), "B2")
        );
        assert_eq!(
            split_address("[Other.xlsx]Data!A:A").unwrap(),
            addr(Some("Other.xlsx"), Some("Data"), "A:A")
        );
        assert_eq!(
            split_address("'[My Book.xlsx]My Sheet'!A1").unwrap(),
            addr(Some("My Book.xlsx"), Some("My Sheet"), "A1")
        );
        assert_eq!(
            split_address("TotalRevenue").unwrap(),
            addr(None, None, "TotalRevenue")
        );
        assert_eq!(
            split_address("  Data!$A$1:$B$2 ").unwrap(),
            addr(None, Some("Data"), "$A$1:$B$2")
        );
        for bad in ["", "   ", "'unterminated!A1", "Data!", "'quoted'A1"] {
            assert!(split_address(bad).is_err(), "{bad:?}");
        }
    }

    #[test]
    fn sheet_names_are_quoted_only_when_needed() {
        assert_eq!(quote_sheet("Data"), "Data");
        assert_eq!(quote_sheet("My Sheet"), "'My Sheet'");
        assert_eq!(quote_sheet("It's"), "'It''s'");
        assert_eq!(quote_sheet("2024"), "'2024'");
    }

    fn bridge(busy_first: u32) -> (Arc<Bridge>, Arc<Mutex<u32>>) {
        let calls = Arc::new(Mutex::new(0u32));
        let c = Arc::clone(&calls);
        let exec: Arc<Executor> = Arc::new(move |op| {
            let mut n = c.lock().unwrap();
            *n += 1;
            if *n <= busy_first {
                return Err(XlError::busy("busy"));
            }
            match op {
                Op::Sheets { .. } => Ok(Reply::Names(vec!["Data".into()])),
                _ => Ok(Reply::Done),
            }
        });
        (
            Arc::new(Bridge::new(exec, Box::new(|| {}), Duration::from_secs(5))),
            calls,
        )
    }

    fn sheets_op() -> Op {
        Op::Sheets {
            workbook: PathBuf::from("/w/b.xlsx"),
        }
    }

    #[test]
    fn the_ui_thread_executes_directly() {
        let (b, calls) = bridge(0);
        assert_eq!(
            b.call(sheets_op()).unwrap(),
            Reply::Names(vec!["Data".into()])
        );
        assert_eq!(*calls.lock().unwrap(), 1);
        assert_eq!(b.queued(), 0);
    }

    #[test]
    fn other_threads_are_served_by_the_ui_thread_pump() {
        let (b, _) = bridge(0);
        let worker = {
            let b = Arc::clone(&b);
            std::thread::spawn(move || b.call(sheets_op()))
        };
        // This (the creating) thread is the "UI thread": keep pumping until the worker is done.
        let start = std::time::Instant::now();
        while !worker.is_finished() {
            b.pump();
            assert!(
                start.elapsed() < Duration::from_secs(5),
                "worker never answered"
            );
            std::thread::sleep(Duration::from_millis(2));
        }
        assert_eq!(
            worker.join().unwrap().unwrap(),
            Reply::Names(vec!["Data".into()])
        );
    }

    #[test]
    fn a_busy_excel_is_retried_and_eventually_gives_up() {
        // Busy twice, then fine.
        let (b, calls) = bridge(2);
        let worker = {
            let b = Arc::clone(&b);
            std::thread::spawn(move || b.call(sheets_op()))
        };
        while !worker.is_finished() {
            b.pump();
            std::thread::sleep(Duration::from_millis(2));
        }
        assert!(worker.join().unwrap().is_ok());
        assert_eq!(*calls.lock().unwrap(), 3);

        // Busy forever: fails after the attempt limit instead of hanging.
        let (b, _) = bridge(u32::MAX);
        let worker = {
            let b = Arc::clone(&b);
            std::thread::spawn(move || b.call(sheets_op()))
        };
        while !worker.is_finished() {
            b.pump();
            std::thread::sleep(Duration::from_millis(1));
        }
        assert!(worker.join().unwrap().unwrap_err().retryable);
    }

    #[test]
    fn a_worker_gives_up_when_the_ui_thread_never_pumps() {
        let exec: Arc<Executor> = Arc::new(|_| Ok(Reply::Done));
        let b = Arc::new(Bridge::new(
            exec,
            Box::new(|| {}),
            Duration::from_millis(150),
        ));
        let b2 = Arc::clone(&b);
        let err = std::thread::spawn(move || b2.call(sheets_op()))
            .join()
            .unwrap()
            .unwrap_err();
        assert!(err.message.contains("did not respond"), "{err}");
    }

    #[test]
    fn read_only_refuses_writes_but_allows_reads() {
        let (b, _) = bridge(0);
        let ro = ReadOnly(b);
        assert!(ro.call(sheets_op()).is_ok());
        let w = Op::Write {
            workbook: PathBuf::new(),
            address: "A1".into(),
            formulas: false,
            data: Value::Num(1.0),
        };
        assert!(ro.call(w).unwrap_err().message.contains("LUA.ASYNC"));
        assert!(
            ro.call(Op::Clear {
                workbook: PathBuf::new(),
                address: "A1".into()
            })
            .is_err()
        );
    }
}
