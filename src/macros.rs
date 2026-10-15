//! Running Lua macros (`xl.macro`) from the ribbon.
//!
//! A macro is a Lua function that is started by the user, not by a cell. It runs on its own
//! thread in a fresh worker state built from the workbook's enabled files, so Excel stays
//! responsive, and it may change cells (through [`crate::xlapi`]). The outcome is queued for the
//! host to report ([`Macros::take_finished`]).

use std::path::{Path, PathBuf};
use std::sync::atomic::AtomicBool;
use std::sync::{Arc, Mutex, OnceLock};
use std::time::{Duration, Instant};

use crate::lua_host::{HostContext, Limits, LuaHost};
use crate::vfs::Bundle;

pub const MACRO_TIMEOUT: Duration = Duration::from_secs(300);

#[derive(Clone, Debug, PartialEq)]
pub struct Finished {
    pub name: String,
    pub workbook: PathBuf,
    pub result: Result<(), String>,
    pub elapsed: Duration,
}

#[derive(Default)]
pub struct Macros {
    running: Mutex<Vec<(PathBuf, String)>>,
    finished: Mutex<Vec<Finished>>,
}

impl Macros {
    /// Starts `name` in the background. Fails if that macro is already running for the workbook.
    pub fn start(
        self: &Arc<Self>,
        workbook: &Path,
        bundle: Arc<Bundle>,
        name: &str,
        cache_dir: Option<PathBuf>,
        timeout: Duration,
    ) -> Result<(), String> {
        let id = (workbook.to_owned(), name.to_owned());
        {
            let mut running = self.running.lock().unwrap();
            if running.contains(&id) {
                return Err(format!("'{name}' is already running"));
            }
            running.push(id.clone());
        }
        let this = Arc::clone(self);
        let spawned = std::thread::Builder::new()
            .name(format!("lua-macro-{name}"))
            .spawn({
                let id = id.clone();
                move || {
                    let started = Instant::now();
                    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                        run(&id.0, &bundle, &id.1, cache_dir.as_deref(), timeout)
                    }))
                    .unwrap_or_else(|_| Err("the macro crashed".to_owned()));
                    this.running.lock().unwrap().retain(|r| *r != id);
                    this.finished.lock().unwrap().push(Finished {
                        name: id.1,
                        workbook: id.0,
                        result,
                        elapsed: started.elapsed(),
                    });
                }
            });
        if let Err(e) = spawned {
            self.running.lock().unwrap().retain(|r| *r != id);
            return Err(format!("cannot start a thread: {e}"));
        }
        Ok(())
    }

    pub fn take_finished(&self) -> Vec<Finished> {
        std::mem::take(&mut *self.finished.lock().unwrap())
    }

    pub fn is_running(&self, workbook: &Path, name: &str) -> bool {
        self.running
            .lock()
            .unwrap()
            .iter()
            .any(|(w, n)| w == workbook && n == name)
    }
}

fn run(
    workbook: &Path,
    bundle: &Bundle,
    name: &str,
    cache_dir: Option<&Path>,
    timeout: Duration,
) -> Result<(), String> {
    let host = LuaHost::new_worker(
        bundle,
        &HostContext {
            workbook,
            cache_dir,
            xl: crate::xlapi::global(),
        },
    )
    .map_err(|e| e.to_string())?;
    let limits = Limits {
        deadline: Instant::now() + timeout,
        cancel: Arc::new(AtomicBool::new(false)),
    };
    host.call_macro(name, &limits).map_err(|e| {
        let m = e.to_string();
        if m.contains("timed out") {
            format!("timed out after {} s", timeout.as_secs())
        } else {
            m
        }
    })
}

pub fn global() -> &'static Arc<Macros> {
    static G: OnceLock<Arc<Macros>> = OnceLock::new();
    G.get_or_init(Arc::default)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::testutil::bundle_of;

    fn wait(m: &Macros) -> Finished {
        let start = Instant::now();
        loop {
            if let Some(f) = m.take_finished().pop() {
                return f;
            }
            assert!(
                start.elapsed() < Duration::from_secs(20),
                "macro never finished"
            );
            std::thread::sleep(Duration::from_millis(5));
        }
    }

    #[test]
    fn a_macro_runs_in_the_background_and_reports() {
        let m = Arc::new(Macros::default());
        let b = Arc::new(bundle_of(&[(
            "t.lua",
            "xl.macro('ok', function() end) xl.macro('bad', function() error('kaput') end) \
             xl.macro('spin', function() while true do end end)",
        )]));
        let wb = Path::new("/w/book.xlsx");
        m.start(wb, Arc::clone(&b), "ok", None, MACRO_TIMEOUT)
            .unwrap();
        let f = wait(&m);
        assert_eq!((f.name.as_str(), f.result), ("ok", Ok(())));
        m.start(wb, Arc::clone(&b), "bad", None, MACRO_TIMEOUT)
            .unwrap();
        assert!(wait(&m).result.unwrap_err().contains("kaput"));
        m.start(wb, Arc::clone(&b), "spin", None, Duration::from_millis(200))
            .unwrap();
        assert!(wait(&m).result.unwrap_err().contains("timed out"));
        assert!(!m.is_running(wb, "spin"));
    }

    #[test]
    fn the_same_macro_cannot_run_twice_at_once() {
        let m = Arc::new(Macros::default());
        let b = Arc::new(bundle_of(&[(
            "t.lua",
            "xl.macro('spin', function() while true do end end)",
        )]));
        let wb = Path::new("/w/book.xlsx");
        m.start(wb, Arc::clone(&b), "spin", None, Duration::from_millis(300))
            .unwrap();
        assert!(m.start(wb, b, "spin", None, MACRO_TIMEOUT).is_err());
        wait(&m);
    }
}
