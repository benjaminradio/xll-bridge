//! Streaming Lua functions behind `LUA.STREAM` / Excel's RTD mechanism.
//!
//! A stream is a Lua function that runs on its own thread for as long as some cell shows it and
//! pushes values with `xl.emit(value)`. Only the latest value matters, so a fast stream cannot
//! build a backlog. Excel decides when a stream starts (`connect`, when a cell first uses it) and
//! ends (`disconnect`, when no cell does); the host collects what changed with
//! [`Streams::take_updates`].
//!
//! Nothing here knows about COM or Excel, so it is tested directly.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use std::time::{Duration, Instant};

use crate::lua_host::{HostContext, Limits, LuaHost};
use crate::value::Value;
use crate::vfs::Bundle;

const MAX_STREAMS: usize = 64;
const MAX_ANNOUNCED: usize = 256;
/// A stream is stopped after this long even if nobody disconnects it.
const LIFETIME: Duration = Duration::from_secs(30 * 24 * 3600);

#[derive(Clone)]
pub struct StreamSpec {
    pub workbook: PathBuf,
    pub bundle: Arc<Bundle>,
    pub func: String,
    pub args: Vec<Value>,
    pub cache_dir: Option<PathBuf>,
}

struct Topic {
    key: String,
    workbook: PathBuf,
    value: Option<Value>,
    dirty: bool,
    cancel: Arc<AtomicBool>,
}

#[derive(Default)]
struct Inner {
    announced: HashMap<String, StreamSpec>,
    topics: HashMap<i32, Topic>,
}

#[derive(Default)]
pub struct Streams {
    inner: Arc<Mutex<Inner>>,
}

impl Streams {
    fn lock(&self) -> std::sync::MutexGuard<'_, Inner> {
        self.inner.lock().unwrap_or_else(|e| e.into_inner())
    }

    /// The worksheet function says "a cell wants this stream". The stream itself starts when
    /// Excel connects a topic with the same `key`. Cheap when already known.
    pub fn announce(&self, key: &str, spec: impl FnOnce() -> StreamSpec) {
        let mut inner = self.lock();
        if let Some((&id, topic)) = inner.topics.iter_mut().find(|(_, t)| t.key == key) {
            // Excel still shows this topic but its stream was stopped (the workbook's scripts were
            // disabled, or re-synced without changing). Excel will not connect it again, so
            // start the stream afresh on the same topic.
            if topic.cancel.load(Ordering::Relaxed) {
                let spec = spec();
                let cancel = Arc::new(AtomicBool::new(false));
                topic.cancel = Arc::clone(&cancel);
                drop(inner);
                if let Err(e) = self.spawn(id, spec, cancel) {
                    crate::log::note(&format!("stream restart failed: {e}"));
                }
            }
            return;
        }
        if inner.announced.contains_key(key) || inner.announced.len() >= MAX_ANNOUNCED {
            return;
        }
        inner.announced.insert(key.to_owned(), spec());
    }

    fn spawn(&self, topic: i32, spec: StreamSpec, cancel: Arc<AtomicBool>) -> Result<(), String> {
        let shared = Arc::clone(&self.inner);
        std::thread::Builder::new()
            .name(format!("lua-stream-{}", spec.func))
            .spawn(move || run(shared, topic, spec, cancel))
            .map(drop)
            .map_err(|e| format!("cannot start a thread: {e}"))
    }

    /// Excel connected `topic` to the stream announced under `key`: start it.
    pub fn connect(&self, topic: i32, key: &str) -> Result<(), String> {
        let (spec, cancel) = {
            let mut inner = self.lock();
            if inner.topics.len() >= MAX_STREAMS {
                return Err(format!("too many streams (the limit is {MAX_STREAMS})"));
            }
            let spec = inner
                .announced
                .remove(key)
                .ok_or("this stream is not known (recalculate the cell)")?;
            let cancel = Arc::new(AtomicBool::new(false));
            inner.topics.insert(
                topic,
                Topic {
                    key: key.to_owned(),
                    workbook: spec.workbook.clone(),
                    value: None,
                    dirty: false,
                    cancel: Arc::clone(&cancel),
                },
            );
            (spec, cancel)
        };
        if let Err(e) = self.spawn(topic, spec, cancel) {
            self.lock().topics.remove(&topic);
            return Err(e);
        }
        Ok(())
    }

    /// Values that changed since the last call, as (topic, value).
    pub fn take_updates(&self) -> Vec<(i32, Value)> {
        let mut inner = self.lock();
        let mut out = Vec::new();
        for (id, t) in inner.topics.iter_mut() {
            if t.dirty
                && let Some(v) = &t.value
            {
                out.push((*id, v.clone()));
            }
            t.dirty = false;
        }
        out
    }

    pub fn has_updates(&self) -> bool {
        self.lock().topics.values().any(|t| t.dirty)
    }

    /// No cell uses this topic any more.
    pub fn disconnect(&self, topic: i32) {
        if let Some(t) = self.lock().topics.remove(&topic) {
            t.cancel.store(true, Ordering::Relaxed);
        }
    }

    /// Stops every stream (the RTD server is shutting down).
    pub fn terminate(&self) {
        let mut inner = self.lock();
        for t in inner.topics.values() {
            t.cancel.store(true, Ordering::Relaxed);
        }
        inner.topics.clear();
        inner.announced.clear();
    }

    /// Stops the streams of one workbook (its scripts were disabled). Excel keeps the topics, and
    /// the cells show the last value until recalculated.
    pub fn cancel_workbook(&self, workbook: &Path) {
        let mut inner = self.lock();
        inner.announced.retain(|_, s| s.workbook != workbook);
        for t in inner.topics.values().filter(|t| t.workbook == workbook) {
            t.cancel.store(true, Ordering::Relaxed);
        }
    }

    pub fn running(&self) -> usize {
        self.lock()
            .topics
            .values()
            .filter(|t| !t.cancel.load(Ordering::Relaxed))
            .count()
    }
}

fn run(shared: Arc<Mutex<Inner>>, topic: i32, spec: StreamSpec, cancel: Arc<AtomicBool>) {
    let publish = {
        let shared = Arc::clone(&shared);
        let cancel = Arc::clone(&cancel);
        move |v: Value| {
            if cancel.load(Ordering::Relaxed) {
                return;
            }
            let mut inner = shared.lock().unwrap_or_else(|e| e.into_inner());
            if let Some(t) = inner.topics.get_mut(&topic) {
                t.value = Some(v);
                t.dirty = true;
            }
        }
    };
    let sink: Arc<dyn Fn(Value) + Send + Sync> = Arc::new(publish.clone());
    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let host = LuaHost::new_worker(
            &spec.bundle,
            &HostContext {
                workbook: &spec.workbook,
                cache_dir: spec.cache_dir.as_deref(),
                xl: crate::xlapi::global(),
            },
        )
        .map_err(|e| e.to_string())?;
        let limits = Limits {
            deadline: Instant::now() + LIFETIME,
            cancel: Arc::clone(&cancel),
        };
        host.run_stream(&spec.func, &spec.args, sink, &limits)
            .map_err(|e| e.to_string())
    }))
    .unwrap_or_else(|_| Err("the stream crashed".to_owned()));
    if let Err(m) = outcome
        && !cancel.load(Ordering::Relaxed)
    {
        publish(Value::Str(format!("#LUA: {m}")));
    }
}

/// A short, stable topic string for (workbook, scripts version, function, arguments): what
/// Excel's RTD identifies a stream by, so identical cells share one stream.
pub fn key_of(workbook: &Path, digest: &str, func: &str, args: &[Value]) -> String {
    use sha2::{Digest, Sha256};
    let mut h = Sha256::new();
    h.update(format!("{}|{digest}|{func}|{args:?}", workbook.display()).as_bytes());
    h.finalize()
        .iter()
        .take(16)
        .map(|b| format!("{b:02x}"))
        .collect()
}

pub fn global() -> &'static Streams {
    static G: OnceLock<Streams> = OnceLock::new();
    G.get_or_init(Streams::default)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::testutil::bundle_of;

    fn spec(code: &str, func: &str, args: Vec<Value>) -> StreamSpec {
        StreamSpec {
            workbook: PathBuf::from("/w/book.xlsx"),
            bundle: Arc::new(bundle_of(&[("t.lua", code)])),
            func: func.into(),
            args,
            cache_dir: None,
        }
    }

    fn wait_for(s: &Streams, f: impl Fn(&[(i32, Value)]) -> bool) -> Vec<(i32, Value)> {
        let start = Instant::now();
        let mut seen = Vec::new();
        while start.elapsed() < Duration::from_secs(20) {
            seen.extend(s.take_updates());
            if f(&seen) {
                return seen;
            }
            std::thread::sleep(Duration::from_millis(5));
        }
        panic!("stream never produced the expected updates: {seen:?}");
    }

    const CODE: &str = "
        function count(n) for i = 1, n do xl.emit(i) xl.sleep(0.01) end end
        function once(x) return x * 2 end
        function boom() xl.emit('before') error('kaput') end
        function forever() local i = 0 while true do i = i + 1 xl.emit(i) xl.sleep(0.01) end end
        function tables() xl.emit({1, 2}) end";

    #[test]
    fn a_stream_pushes_values_and_keeps_only_the_latest() {
        let s = Streams::default();
        s.announce("k", || spec(CODE, "count", vec![Value::Num(5.0)]));
        s.connect(1, "k").unwrap();
        let seen = wait_for(&s, |v| v.iter().any(|(_, x)| *x == Value::Num(5.0)));
        assert!(seen.iter().all(|(t, _)| *t == 1));
        assert!(!s.has_updates());
    }

    #[test]
    fn a_returned_value_is_pushed_too() {
        let s = Streams::default();
        s.announce("k", || spec(CODE, "once", vec![Value::Num(21.0)]));
        s.connect(7, "k").unwrap();
        let seen = wait_for(&s, |v| !v.is_empty());
        assert_eq!(seen.last().unwrap(), &(7, Value::Num(42.0)));
    }

    #[test]
    fn errors_surface_as_text_after_the_values() {
        let s = Streams::default();
        s.announce("k", || spec(CODE, "boom", vec![]));
        s.connect(1, "k").unwrap();
        let seen = wait_for(&s, |v| {
            v.iter()
                .any(|(_, x)| matches!(x, Value::Str(m) if m.contains("kaput")))
        });
        assert!(!seen.is_empty());
        let s2 = Streams::default();
        s2.announce("t", || spec(CODE, "tables", vec![]));
        s2.connect(2, "t").unwrap();
        wait_for(&s2, |v| {
            v.iter()
                .any(|(_, x)| matches!(x, Value::Str(m) if m.contains("single values")))
        });
    }

    #[test]
    fn disconnecting_stops_the_stream() {
        let s = Streams::default();
        s.announce("k", || spec(CODE, "forever", vec![]));
        s.connect(1, "k").unwrap();
        wait_for(&s, |v| !v.is_empty());
        s.disconnect(1);
        std::thread::sleep(Duration::from_millis(150));
        assert!(s.take_updates().is_empty());
        assert_eq!(s.running(), 0);
        // Connecting needs a fresh announcement.
        assert!(s.connect(2, "k").is_err());
    }

    #[test]
    fn a_workbook_can_be_cut_off_and_everything_terminated() {
        let s = Streams::default();
        s.announce("a", || spec(CODE, "forever", vec![]));
        s.connect(1, "a").unwrap();
        s.cancel_workbook(Path::new("/w/book.xlsx"));
        assert_eq!(s.running(), 0);
        s.announce("b", || spec(CODE, "forever", vec![]));
        s.connect(2, "b").unwrap();
        s.terminate();
        assert_eq!(s.running(), 0);
    }

    #[test]
    fn keys_depend_on_everything_that_matters() {
        let wb = Path::new("/w/a.xlsx");
        let k = |d: &str, f: &str, a: Value| key_of(wb, d, f, &[a]);
        assert_eq!(k("d", "f", Value::Num(1.0)), k("d", "f", Value::Num(1.0)));
        assert_ne!(k("d", "f", Value::Num(1.0)), k("d", "f", Value::Num(2.0)));
        assert_ne!(k("d", "f", Value::Num(1.0)), k("e", "f", Value::Num(1.0)));
        assert_ne!(k("d", "f", Value::Num(1.0)), k("d", "g", Value::Num(1.0)));
        assert_eq!(k("d", "f", Value::Empty).len(), 32);
    }

    #[test]
    fn a_stopped_stream_that_excel_still_shows_is_restarted_by_the_next_announcement() {
        let s = Streams::default();
        s.announce("k", || spec(CODE, "forever", vec![]));
        s.connect(1, "k").unwrap();
        wait_for(&s, |v| !v.is_empty());
        s.cancel_workbook(Path::new("/w/book.xlsx"));
        assert_eq!(s.running(), 0);
        std::thread::sleep(Duration::from_millis(100));
        s.take_updates();
        // The cell is recalculated: same topic, so the stream must come back by itself.
        s.announce("k", || spec(CODE, "forever", vec![]));
        assert_eq!(s.running(), 1);
        let seen = wait_for(&s, |v| !v.is_empty());
        assert!(seen.iter().all(|(t, _)| *t == 1));
        s.terminate();
    }

    #[test]
    fn announcing_twice_does_not_replace_a_running_stream() {
        let s = Streams::default();
        s.announce("k", || spec(CODE, "forever", vec![]));
        s.connect(1, "k").unwrap();
        s.announce("k", || panic!("already running"));
        s.disconnect(1);
    }
}
