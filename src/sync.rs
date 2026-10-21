//! Keeping a workbook's embedded files in step with a folder on disk, so scripts can be edited
//! in an external editor (VS Code, ...) and show up in the workbook.
//!
//! The folder is the source of truth: after a sync the workbook holds exactly the folder's files
//! (minus dot-files, editor temp files and anything over the size limits). A [`Watcher`] notices
//! changes by looking at file sizes and modification times, and only reads the files once the
//! folder has been quiet for a poll, so an editor's write-then-rename never produces a half file.
//!
//! Which workbook is linked to which folder is a per-user choice kept in `sync.json`; a workbook
//! can never link itself, so opening someone else's file cannot start reading your disk.

use std::collections::BTreeMap;
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::time::UNIX_EPOCH;

use serde::{Deserialize, Serialize};

use crate::trust::normalize_path;
use crate::vfs::Bundle;

/// Largest number of files/folders looked at; a folder bigger than this is not a script folder.
pub const MAX_ENTRIES: usize = 5000;

pub struct Snapshot {
    pub bundle: Bundle,
    /// Things left out and why (unreadable, too large, ...).
    pub skipped: Vec<String>,
}

/// Names never synced: dot-files/folders (`.git`), lock and temp files of common editors.
fn ignored(name: &str) -> bool {
    name.starts_with('.')
        || name.starts_with("~$")
        || name.ends_with('~')
        || name.ends_with(".tmp")
        || name.ends_with(".swp")
        || name == "node_modules"
}

struct Item {
    rel: String,
    path: PathBuf,
    is_dir: bool,
    len: u64,
    mtime: u128,
}

fn walk(root: &Path) -> Result<Vec<Item>, String> {
    let mut out = Vec::new();
    let mut stack = vec![(root.to_path_buf(), String::new())];
    while let Some((dir, rel)) = stack.pop() {
        let entries = std::fs::read_dir(&dir).map_err(|e| format!("{}: {e}", dir.display()))?;
        for e in entries.filter_map(|e| e.ok()) {
            let name = e.file_name();
            let Some(name) = name.to_str() else { continue };
            if ignored(name) {
                continue;
            }
            let Ok(meta) = std::fs::symlink_metadata(e.path()) else {
                continue;
            };
            if meta.file_type().is_symlink() {
                continue;
            }
            let rel_path = if rel.is_empty() {
                name.to_owned()
            } else {
                format!("{rel}/{name}")
            };
            out.push(Item {
                rel: rel_path.clone(),
                path: e.path(),
                is_dir: meta.is_dir(),
                len: meta.len(),
                mtime: meta
                    .modified()
                    .ok()
                    .and_then(|t| t.duration_since(UNIX_EPOCH).ok())
                    .map_or(0, |d| d.as_nanos()),
            });
            if out.len() > MAX_ENTRIES {
                return Err(format!(
                    "{} has more than {MAX_ENTRIES} files and folders",
                    root.display()
                ));
            }
            if meta.is_dir() {
                stack.push((e.path(), rel_path));
            }
        }
    }
    out.sort_by(|a, b| a.rel.cmp(&b.rel));
    Ok(out)
}

/// A cheap fingerprint of the folder: names, sizes and modification times, no file contents.
pub fn stamp(folder: &Path) -> Result<u64, String> {
    let mut h = DefaultHasher::new();
    for i in walk(folder)? {
        (&i.rel, i.is_dir, i.len, i.mtime).hash(&mut h);
    }
    Ok(h.finish())
}

/// Reads the folder into a [`Bundle`] (text files get LF line endings like everything else).
pub fn scan(folder: &Path) -> Result<Snapshot, String> {
    let mut bundle = Bundle::new();
    let mut skipped = Vec::new();
    for i in walk(folder)? {
        if i.is_dir {
            if let Err(e) = bundle.mkdir(&i.rel) {
                skipped.push(e.to_string());
            }
            continue;
        }
        match std::fs::read(&i.path) {
            Err(e) => skipped.push(format!("{}: {e}", i.rel)),
            Ok(data) => {
                if let Err(e) = bundle.put_file(&i.rel, data, true) {
                    skipped.push(e.to_string());
                }
            }
        }
    }
    Ok(Snapshot { bundle, skipped })
}

pub enum Poll {
    Idle,
    Changed(Snapshot),
    Failed(String),
}

/// Watches one folder. Call [`Watcher::poll`] about once a second.
pub struct Watcher {
    folder: PathBuf,
    applied_stamp: Option<u64>,
    applied_digest: Option<String>,
    pending: Option<u64>,
}

impl Watcher {
    pub fn new(folder: PathBuf) -> Self {
        Watcher {
            folder,
            applied_stamp: None,
            applied_digest: None,
            pending: None,
        }
    }

    pub fn folder(&self) -> &Path {
        &self.folder
    }

    /// Reads the folder now and remembers it as applied (first sync, "sync now").
    pub fn sync_now(&mut self) -> Result<Snapshot, String> {
        let stamp = stamp(&self.folder)?;
        let snap = scan(&self.folder)?;
        self.applied_stamp = Some(stamp);
        self.applied_digest = Some(snap.bundle.digest());
        self.pending = None;
        Ok(snap)
    }

    /// Reports a change once the folder has been quiet (unchanged) for one whole poll.
    pub fn poll(&mut self) -> Poll {
        let now = match stamp(&self.folder) {
            Ok(s) => s,
            Err(e) => {
                // Report a vanished folder once, not every second.
                let first = self.applied_stamp.take().is_some() || self.pending.take().is_some();
                return if first { Poll::Failed(e) } else { Poll::Idle };
            }
        };
        if Some(now) == self.applied_stamp {
            self.pending = None;
            return Poll::Idle;
        }
        if self.pending != Some(now) {
            self.pending = Some(now);
            return Poll::Idle;
        }
        self.pending = None;
        self.applied_stamp = Some(now);
        match scan(&self.folder) {
            Err(e) => Poll::Failed(e),
            Ok(snap) => {
                let digest = snap.bundle.digest();
                if self.applied_digest.as_deref() == Some(digest.as_str()) {
                    Poll::Idle // touched, not changed
                } else {
                    self.applied_digest = Some(digest);
                    Poll::Changed(snap)
                }
            }
        }
    }
}

// ------------------------------------------------------------------ which workbook syncs where

#[derive(Debug, Default, Serialize, Deserialize, PartialEq, Eq)]
pub struct Links {
    /// Normalised workbook path -> folder.
    links: BTreeMap<String, PathBuf>,
}

impl Links {
    pub fn path() -> Option<PathBuf> {
        crate::trust::TrustStore::default_path()
            .and_then(|p| p.parent().map(|d| d.join("sync.json")))
    }

    /// Missing or unreadable means no links.
    pub fn load_from(path: &Path) -> Self {
        std::fs::read(path)
            .ok()
            .and_then(|b| serde_json::from_slice(&b).ok())
            .unwrap_or_default()
    }

    pub fn load() -> Self {
        Self::path()
            .map(|p| Self::load_from(&p))
            .unwrap_or_default()
    }

    /// Writes the file (creating the folder only now that there is something to keep); removes
    /// the file again when the last link is gone.
    pub fn save_to(&self, path: &Path) -> std::io::Result<()> {
        if self.links.is_empty() {
            return match std::fs::remove_file(path) {
                Err(e) if e.kind() != std::io::ErrorKind::NotFound => Err(e),
                _ => Ok(()),
            };
        }
        if let Some(dir) = path.parent() {
            std::fs::create_dir_all(dir)?;
        }
        let json = serde_json::to_vec_pretty(self).map_err(std::io::Error::other)?;
        let tmp = path.with_extension("json.tmp");
        std::fs::write(&tmp, json)?;
        std::fs::rename(tmp, path)
    }

    pub fn save(&self) {
        if let Some(p) = Self::path() {
            let _ = self.save_to(&p);
        }
    }

    pub fn get(&self, workbook: &Path) -> Option<&Path> {
        self.links
            .get(&normalize_path(workbook))
            .map(PathBuf::as_path)
    }

    pub fn set(&mut self, workbook: &Path, folder: PathBuf) {
        self.links.insert(normalize_path(workbook), folder);
    }

    pub fn remove(&mut self, workbook: &Path) -> bool {
        self.links.remove(&normalize_path(workbook)).is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn write(dir: &Path, rel: &str, text: &str) {
        let p = dir.join(rel);
        std::fs::create_dir_all(p.parent().unwrap()).unwrap();
        std::fs::write(p, text).unwrap();
    }

    #[test]
    fn a_folder_becomes_a_bundle_without_editor_litter() {
        let d = tempfile::tempdir().unwrap();
        write(d.path(), "main.lua", "function f() return 1 end\r\n");
        write(d.path(), "lib/util.lua", "return {}");
        write(d.path(), ".git/config", "x");
        write(d.path(), "main.lua~", "backup");
        write(d.path(), "scratch.tmp", "x");
        write(d.path(), "node_modules/a.js", "x");
        std::fs::create_dir_all(d.path().join("empty")).unwrap();
        let snap = scan(d.path()).unwrap();
        let names: Vec<_> = snap
            .bundle
            .files()
            .iter()
            .map(|f| f.path.as_str())
            .collect();
        assert_eq!(names, ["lib/util.lua", "main.lua"]);
        assert_eq!(
            snap.bundle.get("main.lua").unwrap().data,
            b"function f() return 1 end\n"
        );
        assert_eq!(snap.bundle.explicit_dirs(), ["empty", "lib"]);
    }

    #[test]
    fn the_watcher_waits_for_quiet_and_ignores_touches() {
        let d = tempfile::tempdir().unwrap();
        write(d.path(), "main.lua", "a = 1");
        let mut w = Watcher::new(d.path().to_path_buf());
        let first = w.sync_now().unwrap();
        assert_eq!(first.bundle.file_count(), 1);
        assert!(matches!(w.poll(), Poll::Idle));

        // A change is reported on the second poll that sees it, once.
        write(d.path(), "main.lua", "a = 22");
        assert!(matches!(w.poll(), Poll::Idle));
        let Poll::Changed(snap) = w.poll() else {
            panic!("expected a change");
        };
        assert_eq!(snap.bundle.get("main.lua").unwrap().data, b"a = 22");
        assert!(matches!(w.poll(), Poll::Idle));

        // Rewriting identical content changes the stamp but not the content: no event.
        write(d.path(), "main.lua", "a = 22");
        assert!(matches!(w.poll(), Poll::Idle));
        assert!(matches!(w.poll(), Poll::Idle));

        // A new file, then a vanished folder (reported once).
        write(d.path(), "extra.lua", "b = 2");
        assert!(matches!(w.poll(), Poll::Idle));
        assert!(matches!(w.poll(), Poll::Changed(_)));
        let gone = d.path().to_path_buf();
        drop(d);
        assert!(matches!(w.poll(), Poll::Failed(_)));
        assert!(matches!(w.poll(), Poll::Idle));
        assert!(!gone.exists());
    }

    #[test]
    fn links_persist_only_when_there_is_something_to_keep() {
        let d = tempfile::tempdir().unwrap();
        let file = d.path().join("cfg").join("sync.json");
        let mut links = Links::load_from(&file);
        links.save_to(&file).unwrap();
        assert!(
            !file.parent().unwrap().exists(),
            "nothing to keep: nothing created"
        );

        let wb = Path::new("C:/Work/Book.xlsx");
        links.set(wb, PathBuf::from("C:/src/lua"));
        links.save_to(&file).unwrap();
        let again = Links::load_from(&file);
        assert_eq!(
            again.get(Path::new("c:\\work\\book.xlsx")),
            Some(Path::new("C:/src/lua"))
        );

        let mut again = again;
        assert!(again.remove(wb));
        again.save_to(&file).unwrap();
        assert!(!file.exists());
        std::fs::write(&file, b"{ nope").unwrap();
        assert_eq!(Links::load_from(&file), Links::default());
    }
}
