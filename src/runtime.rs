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
use crate::package::{self, ScriptSet};
use crate::trust::{TrustStore, normalize_path};
use crate::value::Value;

pub enum Book {
    NoScripts,
    Blocked {
        digest: String,
        names: Vec<String>,
    },
    Active {
        digest: String,
        names: Vec<String>,
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
    #[error("#UNSAVED: save the workbook to disk first")]
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
        let scripts = match package::read_scripts(workbook) {
            Ok(Some(s)) if !s.is_empty() => s,
            // Unreadable (e.g. mid-save, not yet a zip) is treated like "no scripts"; the next
            // mtime change retries.
            _ => return Book::NoScripts,
        };
        let digest = scripts.digest();
        let names: Vec<String> = scripts.scripts.iter().map(|s| s.name.clone()).collect();
        // A save by Excel changes the mtime but not the scripts: keep the live Lua state.
        if let Some(Book::Active {
            digest: d, host, ..
        }) = previous
            && d == digest
        {
            return Book::Active {
                digest,
                names,
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
        let (digest, scripts) = match package::read_scripts(workbook) {
            Ok(Some(s)) if !s.is_empty() => (s.digest(), s),
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

fn activate(workbook: &Path, scripts: &ScriptSet, digest: String) -> Book {
    match LuaHost::new(scripts, &HostContext { workbook }) {
        Ok(host) => Book::Active {
            digest,
            names: scripts.scripts.iter().map(|s| s.name.clone()).collect(),
            host,
        },
        Err(e) => Book::Failed {
            digest,
            message: e.to_string(),
        },
    }
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
    use crate::package::{Script, embed_scripts, testutil::minimal_xlsx};

    fn workbook_with(code: &str) -> (tempfile::TempDir, PathBuf) {
        let dir = tempfile::tempdir().unwrap();
        let (base, wb) = (dir.path().join("base.xlsx"), dir.path().join("book.xlsx"));
        minimal_xlsx(&base);
        let set = ScriptSet {
            scripts: vec![Script {
                name: "main".into(),
                code: code.into(),
            }],
        };
        embed_scripts(&base, &wb, &set).unwrap();
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
        let set = ScriptSet {
            scripts: vec![Script {
                name: "main".into(),
                code: "function one() return 2 end".into(),
            }],
        };
        embed_scripts(&wb, &edited, &set).unwrap();
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
                names: vec!["main".into()]
            }
        );
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
