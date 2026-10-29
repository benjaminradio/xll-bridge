//! Background execution of Lua calls (`LUA.ASYNC`).
//!
//! A Lua state can only be used from the thread that made it, and the one a workbook's ordinary
//! calls use lives on Excel's UI thread, so a slow call there freezes Excel. Background calls run
//! instead on a small pool of worker threads, each building its own Lua state from the workbook's
//! enabled files (globals set by one state are not visible in another).
//!
//! The worksheet function calls [`Jobs::poll_or_start`]: the first time it gets
//! [`Poll::Pending`] (shown as `#GETTING_DATA`) and a job is queued; once the job is done, the same
//! call returns [`Poll::Ready`]. Finished jobs report which cells were waiting for them
//! ([`Jobs::take_completed`]) so the host can recalculate exactly those cells.
//!
//! Results are cached per (workbook, scripts version, function, arguments) until the arguments or
//! scripts change or [`Jobs::refresh_workbook`] is called. Everything here is plain data plus
//! threads, so it is tested without Excel.

use std::cell::RefCell;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::rc::Rc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex, OnceLock, mpsc};
use std::time::{Duration, Instant};

use crate::lua_host::{HostContext, Limits, LuaHost};
use crate::trust::normalize_path;
use crate::value::Value;
use crate::vfs::Bundle;

pub const DEFAULT_JOB_TIMEOUT: Duration = Duration::from_secs(600);
const WORKERS: usize = 4;
const MAX_ENTRIES: usize = 2000;
const MAX_STATES_PER_WORKER: usize = 8;

/// A worksheet cell (0-based), identified the way COM wants it: workbook name, sheet name, A1.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellRef {
    pub book: String,
    pub sheet: String,
    pub row: u32,
    pub col: u32,
}

impl CellRef {
    /// `A1` style address.
    pub fn a1(&self) -> String {
        let mut letters = Vec::new();
        let mut c = self.col + 1;
        while c > 0 {
            letters.push(b'A' + ((c - 1) % 26) as u8);
            c = (c - 1) / 26;
        }
        letters.reverse();
        format!("{}{}", String::from_utf8_lossy(&letters), self.row + 1)
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct Key {
    workbook: String,
    digest: String,
    func: String,
    args: String,
}

enum State {
    Pending,
    Done(Result<Value, String>),
}

struct Entry {
    id: u64,
    seq: u64,
    state: State,
    callers: Vec<CellRef>,
    cancel: Arc<AtomicBool>,
}

#[derive(Default)]
struct Table {
    entries: HashMap<Key, Entry>,
    completed: Vec<CellRef>,
    next_id: u64,
}

pub struct JobSpec {
    pub workbook: PathBuf,
    pub digest: String,
    pub bundle: Arc<Bundle>,
    pub func: String,
    pub args: Vec<Value>,
    /// The cell to recalculate when the result arrives.
    pub caller: Option<CellRef>,
    pub timeout: Duration,
}

#[derive(Debug, PartialEq)]
pub enum Poll {
    Pending,
    Ready(Result<Value, String>),
}

struct Job {
    key: Key,
    id: u64,
    workbook: PathBuf,
    bundle: Arc<Bundle>,
    func: String,
    args: Vec<Value>,
    timeout: Duration,
    cancel: Arc<AtomicBool>,
    cache_dir: Option<PathBuf>,
    table: Arc<Mutex<Table>>,
}

pub struct Jobs {
    table: Arc<Mutex<Table>>,
    sender: Mutex<Option<mpsc::Sender<Job>>>,
    workers: usize,
    cache_dir: Option<PathBuf>,
}

/// The process-wide pool used by the add-in.
pub fn global() -> &'static Jobs {
    static JOBS: OnceLock<Jobs> = OnceLock::new();
    JOBS.get_or_init(|| Jobs::new(WORKERS, crate::runtime::native_cache_dir()))
}

impl Jobs {
    pub fn new(workers: usize, cache_dir: Option<PathBuf>) -> Self {
        Jobs {
            table: Arc::new(Mutex::new(Table::default())),
            sender: Mutex::new(None),
            workers: workers.max(1),
            cache_dir,
        }
    }

    fn lock(&self) -> std::sync::MutexGuard<'_, Table> {
        // A panic while holding the lock cannot leave the table half-updated in a harmful way.
        self.table.lock().unwrap_or_else(|e| e.into_inner())
    }

    /// Returns the cached outcome, or queues the call (once) and says it is pending.
    pub fn poll_or_start(&self, spec: JobSpec) -> Poll {
        let key = Key {
            workbook: normalize_path(&spec.workbook),
            digest: spec.digest.clone(),
            func: spec.func.clone(),
            args: format!("{:?}", spec.args),
        };
        let job = {
            let mut t = self.lock();
            if let Some(e) = t.entries.get_mut(&key) {
                if let Some(c) = spec.caller
                    && !e.callers.contains(&c)
                {
                    e.callers.push(c);
                }
                return match &e.state {
                    State::Pending => Poll::Pending,
                    State::Done(r) => Poll::Ready(r.clone()),
                };
            }
            evict_if_full(&mut t);
            t.next_id += 1;
            let id = t.next_id;
            let cancel = Arc::new(AtomicBool::new(false));
            t.entries.insert(
                key.clone(),
                Entry {
                    id,
                    seq: id,
                    state: State::Pending,
                    callers: spec.caller.into_iter().collect(),
                    cancel: Arc::clone(&cancel),
                },
            );
            Job {
                key,
                id,
                workbook: spec.workbook,
                bundle: spec.bundle,
                func: spec.func,
                args: spec.args,
                timeout: spec.timeout,
                cancel,
                cache_dir: self.cache_dir.clone(),
                table: Arc::clone(&self.table),
            }
        };
        self.submit(job);
        Poll::Pending
    }

    fn submit(&self, job: Job) {
        let mut sender = self.sender.lock().unwrap_or_else(|e| e.into_inner());
        let tx = sender.get_or_insert_with(|| {
            let (tx, rx) = mpsc::channel::<Job>();
            let rx = Arc::new(Mutex::new(rx));
            for n in 0..self.workers {
                let rx = Arc::clone(&rx);
                let _ = std::thread::Builder::new()
                    .name(format!("lua-worker-{n}"))
                    .spawn(move || worker_loop(rx));
            }
            tx
        });
        let _ = tx.send(job);
    }

    /// Cells whose results have arrived since the last call (each at most once).
    pub fn take_completed(&self) -> Vec<CellRef> {
        let mut done = std::mem::take(&mut self.lock().completed);
        done.sort_by(|a, b| {
            (&a.book, &a.sheet, a.row, a.col).cmp(&(&b.book, &b.sheet, b.row, b.col))
        });
        done.dedup();
        done
    }

    /// Forgets every result for the workbook, cancelling running calls. Returns the cells that
    /// were showing them, so they can be recalculated (which starts the calls afresh).
    pub fn refresh_workbook(&self, workbook: &Path) -> Vec<CellRef> {
        let wb = normalize_path(workbook);
        let mut t = self.lock();
        let mut cells = Vec::new();
        t.entries.retain(|k, e| {
            if k.workbook != wb {
                return true;
            }
            e.cancel.store(true, Ordering::Relaxed);
            cells.extend(e.callers.iter().cloned());
            false
        });
        t.completed.retain(|c| !cells.contains(c));
        cells.sort_by(|a, b| {
            (&a.book, &a.sheet, a.row, a.col).cmp(&(&b.book, &b.sheet, b.row, b.col))
        });
        cells.dedup();
        cells
    }

    /// Excel is shutting down: stop every running call.
    pub fn cancel_all(&self) {
        for e in self.lock().entries.values() {
            e.cancel.store(true, Ordering::Relaxed);
        }
    }

    /// The workbook was closed: drop everything that belonged to it.
    pub fn forget_workbook(&self, workbook: &Path) {
        let _ = self.refresh_workbook(workbook);
    }

    pub fn pending_count(&self) -> usize {
        self.lock()
            .entries
            .values()
            .filter(|e| matches!(e.state, State::Pending))
            .count()
    }
}

fn evict_if_full(t: &mut Table) {
    while t.entries.len() >= MAX_ENTRIES {
        let oldest = t
            .entries
            .iter()
            .filter(|(_, e)| matches!(e.state, State::Done(_)))
            .min_by_key(|(_, e)| e.seq)
            .map(|(k, _)| k.clone());
        match oldest {
            Some(k) => {
                t.entries.remove(&k);
            }
            None => break, // only running calls: let the table grow rather than lose them
        }
    }
}

fn worker_loop(rx: Arc<Mutex<mpsc::Receiver<Job>>>) {
    loop {
        let job = {
            let guard = rx.lock().unwrap_or_else(|e| e.into_inner());
            match guard.recv() {
                Ok(j) => j,
                Err(_) => return, // the pool was dropped
            }
        };
        let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| run_job(&job)))
            .unwrap_or_else(|_| Err("internal error in the worker".to_owned()));
        let mut t = job.table.lock().unwrap_or_else(|e| e.into_inner());
        // A refresh or close in the meantime removed (or replaced) the entry: drop the result.
        let callers = match t.entries.get_mut(&job.key) {
            Some(e) if e.id == job.id => {
                e.state = State::Done(outcome);
                e.callers.clone()
            }
            _ => continue,
        };
        t.completed.extend(callers);
    }
}

type HostKey = (String, String);

thread_local! {
    /// This worker's Lua states, one per (workbook, scripts version).
    static HOSTS: RefCell<Vec<(HostKey, Rc<LuaHost>)>> = const { RefCell::new(Vec::new()) };
}

fn host_for(job: &Job) -> Result<Rc<LuaHost>, String> {
    let key: HostKey = (normalize_path(&job.workbook), job.key.digest.clone());
    if let Some(h) = HOSTS.with(|hs| {
        hs.borrow()
            .iter()
            .find(|(k, _)| *k == key)
            .map(|(_, h)| Rc::clone(h))
    }) {
        return Ok(h);
    }
    let ctx = HostContext {
        workbook: &job.workbook,
        cache_dir: job.cache_dir.as_deref(),
        xl: crate::xlapi::global(),
    };
    let host = Rc::new(LuaHost::new_worker(&job.bundle, &ctx).map_err(|e| e.to_string())?);
    HOSTS.with(|hs| {
        let mut hs = hs.borrow_mut();
        if hs.len() >= MAX_STATES_PER_WORKER {
            hs.remove(0);
        }
        hs.push((key, Rc::clone(&host)));
    });
    Ok(host)
}

fn run_job(job: &Job) -> Result<Value, String> {
    if job.cancel.load(Ordering::Relaxed) {
        return Err("cancelled".to_owned());
    }
    let host = host_for(job)?;
    let limits = Limits {
        deadline: Instant::now() + job.timeout,
        cancel: Arc::clone(&job.cancel),
    };
    host.call_limited(&job.func, &job.args, &limits)
        .map_err(|e| {
            let m = e.to_string();
            if m.contains("timed out") {
                format!("timed out after {} s", job.timeout.as_secs_f64().round())
            } else {
                m
            }
        })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::testutil::bundle_of;

    fn spec(code: &str, func: &str, args: Vec<Value>) -> JobSpec {
        JobSpec {
            workbook: PathBuf::from("/w/book.xlsx"),
            digest: bundle_of(&[("main.lua", code)]).digest(),
            bundle: Arc::new(bundle_of(&[("main.lua", code)])),
            func: func.into(),
            args,
            caller: Some(CellRef {
                book: "book.xlsx".into(),
                sheet: "S".into(),
                row: 0,
                col: 1,
            }),
            timeout: Duration::from_secs(20),
        }
    }

    fn wait_ready(jobs: &Jobs, mk: impl Fn() -> JobSpec) -> Result<Value, String> {
        let start = Instant::now();
        loop {
            if let Poll::Ready(r) = jobs.poll_or_start(mk()) {
                return r;
            }
            assert!(
                start.elapsed() < Duration::from_secs(20),
                "job never finished"
            );
            std::thread::sleep(Duration::from_millis(5));
        }
    }

    const CODE: &str = "
        function add(a, b) return a + b end
        function boom() error('kaput') end
        function spin() while true do end end
        function note(path) local f = io.open(path, 'a') f:write('x\\n') f:close() return 'noted' end
        function slow() local t = os.clock() while os.clock() - t < 0.4 do end return 'slept' end";

    #[test]
    fn cell_addresses() {
        let c = |row, col| {
            CellRef {
                book: "b".into(),
                sheet: "s".into(),
                row,
                col,
            }
            .a1()
        };
        assert_eq!(
            (
                c(0, 0).as_str(),
                c(8, 1).as_str(),
                c(0, 25).as_str(),
                c(9, 26).as_str()
            ),
            ("A1", "B9", "Z1", "AA10")
        );
        assert_eq!(c(1_048_575, 16_383), "XFD1048576");
    }

    #[test]
    fn a_call_is_pending_first_and_then_ready() {
        let jobs = Jobs::new(2, None);
        let mk = || spec(CODE, "slow", vec![]);
        assert_eq!(jobs.poll_or_start(mk()), Poll::Pending);
        assert_eq!(jobs.pending_count(), 1);
        assert_eq!(wait_ready(&jobs, mk).unwrap(), Value::Str("slept".into()));
        assert_eq!(jobs.pending_count(), 0);
    }

    #[test]
    fn the_caller_does_not_wait_for_a_slow_call() {
        let jobs = Jobs::new(2, None);
        let t = Instant::now();
        assert_eq!(
            jobs.poll_or_start(spec(CODE, "slow", vec![])),
            Poll::Pending
        );
        assert!(
            t.elapsed() < Duration::from_millis(300),
            "poll_or_start blocked for {:?}",
            t.elapsed()
        );
    }

    #[test]
    fn results_are_cached_and_each_distinct_call_runs_once() {
        let dir = tempfile::tempdir().unwrap();
        let log = dir.path().join("runs.log");
        let arg = Value::Str(log.to_string_lossy().into_owned());
        let jobs = Jobs::new(2, None);
        let mk = || spec(CODE, "note", vec![arg.clone()]);
        for _ in 0..5 {
            wait_ready(&jobs, mk).unwrap();
        }
        assert_eq!(
            std::fs::read_to_string(&log).unwrap().lines().count(),
            1,
            "ran more than once"
        );

        // Different arguments are a different call.
        assert_eq!(
            wait_ready(&jobs, || spec(
                CODE,
                "add",
                vec![Value::Num(2.0), Value::Num(3.0)]
            ))
            .unwrap(),
            Value::Num(5.0)
        );
        assert_eq!(
            wait_ready(&jobs, || spec(
                CODE,
                "add",
                vec![Value::Num(2.0), Value::Num(4.0)]
            ))
            .unwrap(),
            Value::Num(6.0)
        );
    }

    #[test]
    fn edited_scripts_run_again() {
        let dir = tempfile::tempdir().unwrap();
        let log = dir.path().join("runs.log");
        let arg = Value::Str(log.to_string_lossy().into_owned());
        let jobs = Jobs::new(1, None);
        wait_ready(&jobs, || spec(CODE, "note", vec![arg.clone()])).unwrap();
        // Same function and arguments, but the scripts changed (new digest).
        let edited = format!("{CODE}\n-- edited");
        wait_ready(&jobs, || spec(&edited, "note", vec![arg.clone()])).unwrap();
        assert_eq!(std::fs::read_to_string(&log).unwrap().lines().count(), 2);
    }

    #[test]
    fn errors_come_back_as_results_and_are_cached_too() {
        let jobs = Jobs::new(1, None);
        let r = wait_ready(&jobs, || spec(CODE, "boom", vec![]));
        assert!(r.unwrap_err().contains("kaput"));
        let r = wait_ready(&jobs, || spec(CODE, "missing_function", vec![]));
        assert!(r.unwrap_err().contains("not defined"));
    }

    #[test]
    fn a_runaway_call_times_out() {
        let jobs = Jobs::new(1, None);
        let mk = || JobSpec {
            timeout: Duration::from_millis(300),
            ..spec(CODE, "spin", vec![])
        };
        let r = wait_ready(&jobs, mk);
        assert!(r.unwrap_err().contains("timed out"));
        // The worker survives and serves the next call.
        assert_eq!(
            wait_ready(&jobs, || spec(
                CODE,
                "add",
                vec![Value::Num(1.0), Value::Num(1.0)]
            ))
            .unwrap(),
            Value::Num(2.0)
        );
    }

    #[test]
    fn finished_calls_report_the_cells_that_wait_for_them() {
        let jobs = Jobs::new(2, None);
        let second = CellRef {
            book: "book.xlsx".into(),
            sheet: "S".into(),
            row: 4,
            col: 1,
        };
        let mk = |caller: Option<CellRef>| JobSpec {
            caller,
            ..spec(CODE, "slow", vec![])
        };
        jobs.poll_or_start(mk(None));
        jobs.poll_or_start(mk(Some(second.clone()))); // a second cell with the same call joins the first
        wait_ready(&jobs, || mk(None)).unwrap();
        let done = jobs.take_completed();
        assert!(done.contains(&second), "{done:?}");
        assert!(
            jobs.take_completed().is_empty(),
            "each completion is reported once"
        );
    }

    #[test]
    fn refreshing_discards_results_and_ignores_late_ones() {
        let jobs = Jobs::new(1, None);
        let mk = || spec(CODE, "slow", vec![]);
        jobs.poll_or_start(mk());
        // Refresh while it is still running: its cell comes back for recalculation...
        let cells = jobs.refresh_workbook(Path::new("/w/book.xlsx"));
        assert_eq!(cells.len(), 1);
        // ...and the abandoned run's result must not resurrect the entry.
        std::thread::sleep(Duration::from_millis(700));
        assert!(jobs.take_completed().is_empty());
        assert_eq!(jobs.pending_count(), 0);
        // Asking again starts a fresh call.
        assert_eq!(wait_ready(&jobs, mk).unwrap(), Value::Str("slept".into()));
    }

    #[test]
    fn workbooks_are_kept_apart() {
        let jobs = Jobs::new(1, None);
        wait_ready(&jobs, || {
            spec(CODE, "add", vec![Value::Num(1.0), Value::Num(2.0)])
        })
        .unwrap();
        let other = JobSpec {
            workbook: PathBuf::from("/other/book2.xlsx"),
            ..spec(CODE, "add", vec![Value::Num(1.0), Value::Num(2.0)])
        };
        // Same function, same arguments, different workbook: not a cache hit.
        assert_eq!(jobs.poll_or_start(other), Poll::Pending);
        jobs.forget_workbook(Path::new("/w/book.xlsx"));
        assert_eq!(jobs.refresh_workbook(Path::new("/w/book.xlsx")).len(), 0);
    }
}
