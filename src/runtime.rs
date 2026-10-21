//! Per-workbook state machine tying storage, trust and the Lua host together.
//!
//! ```text
//!   no scripts part ──────────────► NoScripts
//!   scripts, not trusted ─────────► Blocked  ──enable()──► Active
//!   scripts, trusted ─────────────► Active
//!   Lua failed to load ───────────► Failed
//! ```

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::time::SystemTime;

use crate::lua_host::{HostContext, HostError, LuaHost};
use std::sync::Arc;

use crate::package;
use crate::trust::{TrustStore, normalize_path};
use crate::value::Value;
use crate::vfs::Bundle;

pub enum Book {
    NoScripts,
    Blocked {
        digest: String,
        names: Vec<String>,
    },
    Active {
        digest: String,
        names: Vec<String>,
        bundle: Arc<Bundle>,
        host: LuaHost,
    },
    Failed {
        digest: String,
        message: String,
    },
}

/// What a UI needs to know about a workbook, without access to the Lua state.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Status {
    /// No scripts (or no file on disk): nothing to trust.
    NoScripts,
    /// Scripts present and trusted; Lua is running.
    Active { names: Vec<String> },
    /// Scripts present but not trusted; nothing has been executed.
    Blocked { names: Vec<String> },
    /// Trusted, but loading the scripts failed.
    Failed { message: String },
}

struct Entry {
    mtime: Option<SystemTime>,
    book: Book,
}

/// Why a call produced no value. `Display` is what ends up in the cell.
#[derive(Debug, thiserror::Error)]
pub enum CallError {
    #[error("#BLOCKED: Lua scripts in this workbook are disabled (use the Enable button)")]
    Blocked,
    #[error("#NOSCRIPTS: this workbook has no embedded Lua scripts")]
    NoScripts,
    #[error("#NOPATH: cannot locate this workbook on disk")]
    NoPath,
    #[error("#LUA: {0}")]
    Lua(String),
}

impl From<HostError> for CallError {
    fn from(e: HostError) -> Self {
        CallError::Lua(e.to_string())
    }
}

pub struct Runtime {
    trust: TrustStore,
    books: HashMap<String, Entry>,
    /// Shared sandbox for `LUA.EVAL` where there is no trusted script state.
    sandbox: Option<LuaHost>,
}

impl Runtime {
    pub fn new(trust: TrustStore) -> Self {
        Runtime {
            trust,
            books: HashMap::new(),
            sandbox: None,
        }
    }

    /// Current state for a workbook, (re)reading it from disk if the file changed.
    pub fn state(&mut self, workbook: &Path) -> &Book {
        let key = normalize_path(workbook);
        let mtime = std::fs::metadata(workbook).and_then(|m| m.modified()).ok();
        let stale = self.books.get(&key).is_none_or(|e| e.mtime != mtime);
        if stale {
            let old = self.books.remove(&key);
            let book = self.load(workbook, old.map(|e| e.book));
            self.books.insert(key.clone(), Entry { mtime, book });
        }
        &self.books[&key].book
    }

    fn load(&mut self, workbook: &Path, previous: Option<Book>) -> Book {
        let scripts = match package::read_bundle(workbook) {
            Ok(Some(s)) if s.file_count() > 0 => s,
            // Unreadable (e.g. mid-save, not yet a zip) is treated like "no scripts"; the next
            // mtime change retries.
            _ => return Book::NoScripts,
        };
        let digest = scripts.digest();
        let names = script_names(&scripts);
        // A save by Excel changes the mtime but not the scripts: keep the live Lua state.
        if let Some(Book::Active {
            digest: d,
            host,
            bundle,
            ..
        }) = previous
            && d == digest
        {
            return Book::Active {
                digest,
                names,
                bundle,
                host,
            };
        }
        if !self.trust.is_trusted(workbook, &digest) {
            return Book::Blocked { digest, names };
        }
        activate(workbook, &scripts, digest)
    }

    /// User clicked "Enable": trust the scripts as currently on disk and start Lua.
    pub fn enable(&mut self, workbook: &Path) -> std::io::Result<()> {
        let key = normalize_path(workbook);
        let (digest, scripts) = match package::read_bundle(workbook) {
            Ok(Some(s)) if s.file_count() > 0 => (s.digest(), s),
            _ => return Err(std::io::Error::other("no scripts to enable")),
        };
        self.trust.trust(workbook, &digest)?;
        let mtime = std::fs::metadata(workbook).and_then(|m| m.modified()).ok();
        self.books.insert(
            key,
            Entry {
                mtime,
                book: activate(workbook, &scripts, digest),
            },
        );
        Ok(())
    }

    /// Takes back the user's permission: the workbook's scripts stop running and are blocked
    /// again until enabled anew.
    pub fn revoke(&mut self, workbook: &Path) -> std::io::Result<()> {
        self.trust.revoke(workbook)?;
        // Dropping the cached state drops the live Lua state; the reload finds no trust.
        self.books.remove(&normalize_path(workbook));
        Ok(())
    }

    pub fn call(
        &mut self,
        workbook: Option<&Path>,
        func: &str,
        args: &[Value],
    ) -> Result<Value, CallError> {
        Ok(self.active(workbook)?.call(func, args)?)
    }

    /// Evaluates an expression. With no trusted scripts to run in (unsaved workbook, or one
    /// without scripts) it runs in the sandbox; a workbook whose scripts are blocked stays blocked.
    pub fn eval(
        &mut self,
        workbook: Option<&Path>,
        code: &str,
        args: &[Value],
    ) -> Result<Value, CallError> {
        let sandboxed = match workbook {
            None => true,
            Some(p) => matches!(self.state(p), Book::NoScripts),
        };
        if sandboxed {
            if self.sandbox.is_none() {
                self.sandbox = Some(LuaHost::sandboxed()?);
            }
            let host = self.sandbox.as_ref().expect("just created");
            return Ok(host.eval(code, args)?);
        }
        Ok(self.active(workbook)?.eval(code, args)?)
    }

    fn active(&mut self, workbook: Option<&Path>) -> Result<&LuaHost, CallError> {
        let wb = workbook.ok_or(CallError::NoPath)?;
        match self.state(wb) {
            Book::Active { host, .. } => Ok(host),
            Book::Blocked { .. } => Err(CallError::Blocked),
            Book::NoScripts => Err(CallError::NoScripts),
            Book::Failed { message, .. } => Err(CallError::Lua(message.clone())),
        }
    }

    /// The files and digest of an enabled workbook, for background workers (which build their own
    /// Lua state from them). Errors exactly like a synchronous call would.
    pub fn enabled_files(
        &mut self,
        workbook: Option<&Path>,
    ) -> Result<(Arc<Bundle>, String), CallError> {
        let wb = workbook.ok_or(CallError::NoPath)?;
        match self.state(wb) {
            Book::Active { bundle, digest, .. } => Ok((Arc::clone(bundle), digest.clone())),
            Book::Blocked { .. } => Err(CallError::Blocked),
            Book::NoScripts => Err(CallError::NoScripts),
            Book::Failed { message, .. } => Err(CallError::Lua(message.clone())),
        }
    }

    /// Macros (`xl.macro`) defined by an enabled workbook; empty when it is not enabled.
    pub fn macro_names(&mut self, workbook: &Path) -> Vec<String> {
        match self.state(workbook) {
            Book::Active { host, .. } => host.macro_names(),
            _ => Vec::new(),
        }
    }

    /// Digest of the file set on disk for this workbook (None when there are no files).
    pub fn disk_digest(&mut self, workbook: &Path) -> Option<String> {
        match self.state(workbook) {
            Book::NoScripts => None,
            Book::Blocked { digest, .. }
            | Book::Active { digest, .. }
            | Book::Failed { digest, .. } => Some(digest.clone()),
        }
    }

    /// The user changed this workbook's embedded files through the panel (not yet saved).
    /// If the workbook had no files, or was already trusted, the author is the user: trust the
    /// new file set too so it stays enabled once saved. A blocked workbook stays blocked.
    /// Returns whether the new set was trusted.
    pub fn note_edit(&mut self, workbook: &Path, new_digest: &str) -> bool {
        let allowed = matches!(self.state(workbook), Book::NoScripts | Book::Active { .. });
        allowed && self.trust.pretrust(workbook, new_digest).is_ok()
    }

    /// Makes `bundle` the trusted, running file set of `workbook` right away, without waiting for
    /// the workbook to be saved. For a folder the user linked on purpose (see `sync`): the files
    /// are theirs. Saving later keeps the running state, since the saved digest matches.
    pub fn apply_live(&mut self, workbook: &Path, bundle: &Bundle) -> std::io::Result<()> {
        if bundle.file_count() == 0 {
            return Err(std::io::Error::other("there are no files to run"));
        }
        let digest = bundle.digest();
        self.trust.pretrust(workbook, &digest)?;
        let mtime = std::fs::metadata(workbook).and_then(|m| m.modified()).ok();
        self.books.insert(
            normalize_path(workbook),
            Entry {
                mtime,
                book: activate(workbook, bundle, digest),
            },
        );
        Ok(())
    }

    /// A file set the user authored in an unsaved workbook has now been saved to `workbook`:
    /// trust it (if the saved digest is the one they authored) and reload so it activates.
    pub fn adopt_authored(&mut self, workbook: &Path, digest: &str) -> bool {
        if self.disk_digest(workbook).as_deref() != Some(digest)
            || self.trust.pretrust(workbook, digest).is_err()
        {
            return false;
        }
        self.books.remove(&normalize_path(workbook));
        true
    }

    /// Summary for UI (ribbon, `LUA.STATUS`).
    pub fn status(&mut self, workbook: &Path) -> Status {
        match self.state(workbook) {
            Book::NoScripts => Status::NoScripts,
            Book::Blocked { names, .. } => Status::Blocked {
                names: names.clone(),
            },
            Book::Active { names, .. } => Status::Active {
                names: names.clone(),
            },
            Book::Failed { message, .. } => Status::Failed {
                message: message.clone(),
            },
        }
    }
}

/// Names shown in the UI: every embedded file, capped.
fn script_names(bundle: &Bundle) -> Vec<String> {
    const CAP: usize = 40;
    let mut names: Vec<String> = bundle
        .entries()
        .into_iter()
        .filter(|e| e.kind == crate::vfs::Kind::File)
        .map(|e| e.path)
        .collect();
    if names.len() > CAP {
        let more = names.len() - CAP;
        names.truncate(CAP);
        names.push(format!("\u{2026} and {more} more"));
    }
    names
}

/// `%LOCALAPPDATA%\xll-bridge\cache` (or the XDG/HOME equivalent): extracted native modules.
pub fn native_cache_dir() -> Option<PathBuf> {
    let base = std::env::var_os("LOCALAPPDATA")
        .map(PathBuf::from)
        .or_else(|| std::env::var_os("XDG_CACHE_HOME").map(PathBuf::from))
        .or_else(|| std::env::var_os("HOME").map(|h| PathBuf::from(h).join(".cache")))?;
    Some(base.join("xll-bridge").join("cache"))
}

fn activate(workbook: &Path, bundle: &Bundle, digest: String) -> Book {
    let cache = native_cache_dir();
    match LuaHost::new(
        bundle,
        &HostContext {
            workbook,
            cache_dir: cache.as_deref(),
            xl: crate::xlapi::global().map(|a| {
                std::sync::Arc::new(crate::xlapi::ReadOnly(a))
                    as std::sync::Arc<dyn crate::xlapi::XlApi>
            }),
        },
    ) {
        Ok(host) => Book::Active {
            digest,
            names: script_names(bundle),
            bundle: Arc::new(bundle.clone()),
            host,
        },
        Err(e) => Book::Failed {
            digest,
            message: e.to_string(),
        },
    }
}

/// `"[demo.xlsx]Sheet1"` (what `xlSheetNm` returns) -> `"demo.xlsx"`.
pub fn workbook_name_from_sheet_name(s: &str) -> Option<String> {
    let open = s.find('[')?;
    let close = s[open..].find(']')? + open;
    let name = &s[open + 1..close];
    (!name.is_empty()).then(|| name.to_owned())
}

/// Splits Excel's `CELL("filename")` result, `C:\dir\[Book.xlsx]Sheet1`, into the file path.
/// `None` for unsaved workbooks (empty string) and non-local locations (URLs).
pub fn workbook_path_from_cell_filename(s: &str) -> Option<PathBuf> {
    let open = s.find('[')?;
    let close = s[open..].find(']')? + open;
    let (dir, name) = (&s[..open], &s[open + 1..close]);
    if name.is_empty() || dir.contains("://") {
        return None;
    }
    Some(PathBuf::from(format!("{dir}{name}")))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::{
        embed_bundle,
        testutil::{bundle_of, minimal_xlsx},
    };

    fn workbook_with(code: &str) -> (tempfile::TempDir, PathBuf) {
        let dir = tempfile::tempdir().unwrap();
        let (base, wb) = (dir.path().join("base.xlsx"), dir.path().join("book.xlsx"));
        minimal_xlsx(&base);
        let set = bundle_of(&[("main.lua", code)]);
        embed_bundle(&base, &wb, &set).unwrap();
        (dir, wb)
    }

    #[test]
    fn blocked_until_enabled_and_never_executes_before() {
        // The script has a visible side effect at load time: it must not happen while blocked.
        let marker = tempfile::tempdir().unwrap().keep().join("ran");
        let code = format!(
            "local f = io.open([[{}]], 'w') f:write('x') f:close() function one() return 1 end",
            marker.display()
        );
        let (_d, wb) = workbook_with(&code);
        let mut rt = Runtime::new(TrustStore::ephemeral());

        assert!(matches!(
            rt.call(Some(&wb), "one", &[]),
            Err(CallError::Blocked)
        ));
        assert!(matches!(rt.status(&wb), Status::Blocked { .. }));
        assert!(!marker.exists(), "Lua ran while the workbook was untrusted");

        rt.enable(&wb).unwrap();
        assert_eq!(rt.call(Some(&wb), "one", &[]).unwrap(), Value::Num(1.0));
        assert!(marker.exists());
    }

    #[test]
    fn editing_scripts_revokes_trust() {
        let (d, wb) = workbook_with("function one() return 1 end");
        let mut rt = Runtime::new(TrustStore::ephemeral());
        rt.enable(&wb).unwrap();
        assert!(rt.call(Some(&wb), "one", &[]).is_ok());

        let edited = d.path().join("edited.xlsx");
        let set = bundle_of(&[("main.lua", "function one() return 2 end")]);
        embed_bundle(&wb, &edited, &set).unwrap();
        std::fs::copy(&edited, &wb).unwrap();
        // mtime granularity can hide the change on fast filesystems; force a fresh read.
        rt.books.clear();
        assert!(matches!(
            rt.call(Some(&wb), "one", &[]),
            Err(CallError::Blocked)
        ));
    }

    #[test]
    fn eval_is_sandboxed_without_scripts_but_blocked_with_untrusted_ones() {
        let dir = tempfile::tempdir().unwrap();
        let plain = dir.path().join("plain.xlsx");
        minimal_xlsx(&plain);
        let mut rt = Runtime::new(TrustStore::ephemeral());
        // Unsaved and script-less workbooks may evaluate expressions, with no ambient authority.
        assert_eq!(rt.eval(None, "1 +2", &[]).unwrap(), Value::Num(3.0));
        assert_eq!(rt.eval(Some(&plain), "1 +2", &[]).unwrap(), Value::Num(3.0));
        assert_eq!(rt.eval(None, "io == nil", &[]).unwrap(), Value::Bool(true));

        let (_d, wb) = workbook_with("function f() return 1 end");
        assert!(matches!(
            rt.eval(Some(&wb), "1", &[]),
            Err(CallError::Blocked)
        ));
        assert_eq!(rt.status(&plain), Status::NoScripts);
        rt.enable(&wb).unwrap();
        assert_eq!(
            rt.eval(Some(&wb), "io ~= nil", &[]).unwrap(),
            Value::Bool(true)
        );
        assert_eq!(
            rt.status(&wb),
            Status::Active {
                names: vec!["main.lua".into()]
            }
        );
    }

    #[test]
    fn editing_through_the_panel_keeps_trusted_workbooks_trusted_but_not_blocked_ones() {
        let (d, wb) = workbook_with("function one() return 1 end");
        let mut rt = Runtime::new(TrustStore::ephemeral());
        // Blocked: an edit must not grant trust.
        assert!(!rt.note_edit(&wb, "new-digest"));
        rt.enable(&wb).unwrap();
        // Trusted: the author's edit is pre-trusted; the digest on disk stays trusted too.
        let disk = rt.disk_digest(&wb).unwrap();
        assert!(rt.note_edit(&wb, "new-digest"));
        assert!(rt.trust.is_trusted(&wb, "new-digest") && rt.trust.is_trusted(&wb, &disk));
        // Saved-as later: the digest they authored while unsaved is adopted, others are not.
        assert!(!rt.adopt_authored(&wb, "something-else"));
        let (_d2, wb2) = workbook_with("function two() return 2 end");
        let authored = rt.disk_digest(&wb2).unwrap();
        assert!(matches!(rt.status(&wb2), Status::Blocked { .. }));
        assert!(rt.adopt_authored(&wb2, &authored));
        assert!(matches!(rt.status(&wb2), Status::Active { .. }));
        // A workbook without files: the first files the user adds are theirs.
        let plain = d.path().join("plain.xlsx");
        minimal_xlsx(&plain);
        assert_eq!(rt.disk_digest(&plain), None);
        assert!(rt.note_edit(&plain, "first"));
    }

    #[test]
    fn only_empty_folders_do_not_count_as_scripts() {
        let dir = tempfile::tempdir().unwrap();
        let (base, wb) = (dir.path().join("b.xlsx"), dir.path().join("w.xlsx"));
        minimal_xlsx(&base);
        let mut b = crate::vfs::Bundle::new();
        b.mkdir("only/folders").unwrap();
        embed_bundle(&base, &wb, &b).unwrap();
        let mut rt = Runtime::new(TrustStore::ephemeral());
        assert_eq!(rt.status(&wb), Status::NoScripts);
    }

    #[test]
    fn the_shipped_demo_runs_end_to_end_with_a_module_in_a_folder() {
        let dir = tempfile::tempdir().unwrap();
        let (base, wb) = (dir.path().join("b.xlsx"), dir.path().join("demo.xlsx"));
        minimal_xlsx(&base);
        let b = bundle_of(&[
            ("demo.lua", include_str!("../examples/demo.lua")),
            ("lib/mathx.lua", include_str!("../examples/lib/mathx.lua")),
        ]);
        embed_bundle(&base, &wb, &b).unwrap();
        let mut rt = Runtime::new(TrustStore::ephemeral());
        assert!(matches!(
            rt.call(Some(&wb), "hypot", &[]),
            Err(CallError::Blocked)
        ));
        rt.enable(&wb).unwrap();
        let args = [Value::Num(3.0), Value::Num(4.0)];
        assert_eq!(rt.call(Some(&wb), "hypot", &args).unwrap(), Value::Num(5.0));
        assert_eq!(rt.call(Some(&wb), "add", &args).unwrap(), Value::Num(7.0));
    }

    #[test]
    fn workers_get_the_files_of_an_enabled_workbook_only() {
        let (_d, wb) = workbook_with("function one() return 1 end");
        let mut rt = Runtime::new(TrustStore::ephemeral());
        assert!(matches!(
            rt.enabled_files(Some(&wb)),
            Err(CallError::Blocked)
        ));
        rt.enable(&wb).unwrap();
        let (files, digest) = rt.enabled_files(Some(&wb)).unwrap();
        assert_eq!(files.digest(), digest);
        assert!(files.get("main.lua").is_some());
        assert!(matches!(rt.enabled_files(None), Err(CallError::NoPath)));

        // The files handed over really run in a background worker, with a module from a folder.
        let jobs = crate::jobs::Jobs::new(1, None);
        let spec = || crate::jobs::JobSpec {
            workbook: wb.clone(),
            digest: digest.clone(),
            bundle: Arc::clone(&files),
            func: "one".into(),
            args: vec![],
            caller: None,
            timeout: std::time::Duration::from_secs(20),
        };
        let start = std::time::Instant::now();
        let result = loop {
            if let crate::jobs::Poll::Ready(r) = jobs.poll_or_start(spec()) {
                break r;
            }
            assert!(start.elapsed().as_secs() < 20, "job never finished");
            std::thread::sleep(std::time::Duration::from_millis(5));
        };
        assert_eq!(result, Ok(Value::Num(1.0)));
    }

    #[test]
    fn revoking_blocks_a_trusted_workbook_again() {
        let (_d, wb) = workbook_with("function one() return 1 end");
        let mut rt = Runtime::new(TrustStore::ephemeral());
        rt.enable(&wb).unwrap();
        assert_eq!(rt.call(Some(&wb), "one", &[]).unwrap(), Value::Num(1.0));
        rt.revoke(&wb).unwrap();
        assert!(matches!(
            rt.call(Some(&wb), "one", &[]),
            Err(CallError::Blocked)
        ));
        assert!(matches!(rt.status(&wb), Status::Blocked { .. }));
        rt.enable(&wb).unwrap();
        assert_eq!(rt.call(Some(&wb), "one", &[]).unwrap(), Value::Num(1.0));
    }

    #[test]
    fn a_live_file_set_runs_before_the_workbook_is_saved() {
        let (_d, wb) = workbook_with("function one() return 1 end");
        let mut rt = Runtime::new(TrustStore::ephemeral());
        // Not trusted yet, and the new file set is not on disk.
        let live =
            crate::package::testutil::bundle_of(&[("main.lua", "function one() return 2 end")]);
        rt.apply_live(&wb, &live).unwrap();
        assert_eq!(rt.call(Some(&wb), "one", &[]).unwrap(), Value::Num(2.0));
        assert!(matches!(rt.status(&wb), Status::Active { .. }));
        assert!(rt.apply_live(&wb, &Bundle::new()).is_err());
    }

    #[test]
    fn no_scripts_and_no_path() {
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("plain.xlsx");
        minimal_xlsx(&wb);
        let mut rt = Runtime::new(TrustStore::ephemeral());
        assert!(matches!(
            rt.call(Some(&wb), "f", &[]),
            Err(CallError::NoScripts)
        ));
        assert!(matches!(rt.call(None, "f", &[]), Err(CallError::NoPath)));
    }

    #[test]
    fn extracts_the_workbook_name_from_a_sheet_reference() {
        assert_eq!(
            workbook_name_from_sheet_name("[demo.xlsx]Demo").as_deref(),
            Some("demo.xlsx")
        );
        assert_eq!(
            workbook_name_from_sheet_name("[Book 1]Sheet1").as_deref(),
            Some("Book 1")
        );
        assert_eq!(workbook_name_from_sheet_name("Sheet1"), None);
        assert_eq!(workbook_name_from_sheet_name("[]Sheet1"), None);
    }

    #[test]
    fn parses_cell_filename() {
        assert_eq!(
            workbook_path_from_cell_filename(r"C:\Users\me\[Book1.xlsx]Sheet1"),
            Some(PathBuf::from(r"C:\Users\me\Book1.xlsx"))
        );
        assert_eq!(workbook_path_from_cell_filename(""), None);
        assert_eq!(
            workbook_path_from_cell_filename("https://x/y/[a.xlsx]S"),
            None
        );
    }
}
