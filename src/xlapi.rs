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

impl Op {
    pub fn is_write(&self) -> bool {
        matches!(self, Op::Write { .. } | Op::Clear { .. })
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
            })
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
