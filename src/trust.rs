//! Trust store: Lua in a workbook never runs until the user enables it.
//!
//! Trust is pinned to (workbook path, script digest). Editing any script, or moving the file,
//! invalidates it and the banner comes back — the same shape as VBA's trusted documents.

use std::path::{Path, PathBuf};

use serde::{Deserialize, Serialize};

const MAX_PER_WORKBOOK: usize = 4;

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Eq)]
struct Entry {
    workbook: String,
    digest: String,
}

#[derive(Debug, Default, Serialize, Deserialize)]
struct Disk {
    trusted: Vec<Entry>,
}

#[derive(Debug)]
pub struct TrustStore {
    path: Option<PathBuf>,
    entries: Vec<Entry>,
}

/// Case-insensitive, separator-normalised key: Windows paths compare that way.
pub fn normalize_path(p: &Path) -> String {
    p.to_string_lossy().replace('/', "\\").to_lowercase()
}

impl TrustStore {
    /// In-memory only; nothing persists. Used by tests.
    pub fn ephemeral() -> Self {
        TrustStore {
            path: None,
            entries: Vec::new(),
        }
    }

    /// `%APPDATA%\xll-bridge\trust.json` (or `$XDG_CONFIG_HOME`/`~/.config` off Windows).
    pub fn default_path() -> Option<PathBuf> {
        let base = std::env::var_os("APPDATA")
            .map(PathBuf::from)
            .or_else(|| std::env::var_os("XDG_CONFIG_HOME").map(PathBuf::from))
            .or_else(|| std::env::var_os("HOME").map(|h| PathBuf::from(h).join(".config")))?;
        Some(base.join("xll-bridge").join("trust.json"))
    }

    /// A missing or unreadable file means nothing is trusted (fail closed).
    pub fn load(path: PathBuf) -> Self {
        let entries = std::fs::read(&path)
            .ok()
            .and_then(|b| serde_json::from_slice::<Disk>(&b).ok())
            .map(|d| d.trusted)
            .unwrap_or_default();
        TrustStore {
            path: Some(path),
            entries,
        }
    }

    pub fn is_trusted(&self, workbook: &Path, digest: &str) -> bool {
        let key = normalize_path(workbook);
        self.entries
            .iter()
            .any(|e| e.workbook == key && e.digest == digest)
    }

    /// The user explicitly enabled this exact file set: it becomes the only trusted one.
    pub fn trust(&mut self, workbook: &Path, digest: &str) -> std::io::Result<()> {
        let key = normalize_path(workbook);
        self.entries.retain(|e| e.workbook != key);
        self.entries.push(Entry {
            workbook: key,
            digest: digest.to_owned(),
        });
        self.save()
    }

    /// The user edited a trusted (or file-free) workbook themselves; also trust what they have
    /// built so it stays enabled once saved, without revoking the version currently on disk.
    /// At most [`MAX_PER_WORKBOOK`] digests are kept per workbook (oldest dropped).
    pub fn pretrust(&mut self, workbook: &Path, digest: &str) -> std::io::Result<()> {
        let key = normalize_path(workbook);
        if self
            .entries
            .iter()
            .any(|e| e.workbook == key && e.digest == digest)
        {
            return Ok(());
        }
        self.entries.push(Entry {
            workbook: key.clone(),
            digest: digest.to_owned(),
        });
        let mut excess = self
            .entries
            .iter()
            .filter(|e| e.workbook == key)
            .count()
            .saturating_sub(MAX_PER_WORKBOOK);
        self.entries.retain(|e| {
            if excess > 0 && e.workbook == key {
                excess -= 1;
                false
            } else {
                true
            }
        });
        self.save()
    }

    /// Forgets every trusted file set of this workbook; its scripts are blocked again.
    pub fn revoke(&mut self, workbook: &Path) -> std::io::Result<()> {
        let key = normalize_path(workbook);
        let before = self.entries.len();
        self.entries.retain(|e| e.workbook != key);
        if self.entries.len() == before {
            return Ok(());
        }
        self.save()
    }

    fn save(&self) -> std::io::Result<()> {
        let Some(path) = &self.path else {
            return Ok(());
        };
        if let Some(dir) = path.parent() {
            std::fs::create_dir_all(dir)?;
        }
        let json = serde_json::to_vec_pretty(&Disk {
            trusted: self.entries.clone(),
        })
        .map_err(std::io::Error::other)?;
        // Write-then-rename so a crash can't leave a truncated (hence empty = untrusted) file.
        let tmp = path.with_extension("json.tmp");
        std::fs::write(&tmp, json)?;
        std::fs::rename(tmp, path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn trust_is_pinned_to_digest_and_path() {
        let mut t = TrustStore::ephemeral();
        let wb = Path::new(r"C:\Work\Book.xlsx");
        assert!(!t.is_trusted(wb, "aa"));
        t.trust(wb, "aa").unwrap();
        assert!(t.is_trusted(Path::new(r"c:/work/book.XLSX"), "aa"));
        assert!(
            !t.is_trusted(wb, "bb"),
            "edited scripts must not stay trusted"
        );
        assert!(!t.is_trusted(Path::new(r"C:\Other\Book.xlsx"), "aa"));
        t.trust(wb, "bb").unwrap();
        assert!(!t.is_trusted(wb, "aa"), "new digest replaces the old one");
    }

    #[test]
    fn revoke_forgets_only_that_workbook() {
        let mut t = TrustStore::ephemeral();
        let (a, b) = (Path::new("/w/a.xlsx"), Path::new("/w/b.xlsx"));
        t.trust(a, "d1").unwrap();
        t.pretrust(a, "d2").unwrap();
        t.trust(b, "d3").unwrap();
        t.revoke(a).unwrap();
        assert!(!t.is_trusted(a, "d1") && !t.is_trusted(a, "d2"));
        assert!(t.is_trusted(b, "d3"));
        t.revoke(a).unwrap(); // nothing left: fine
    }

    #[test]
    fn pretrust_adds_without_revoking_and_is_bounded() {
        let mut t = TrustStore::ephemeral();
        let wb = Path::new("/w/a.xlsx");
        t.trust(wb, "disk").unwrap();
        t.pretrust(wb, "edit1").unwrap();
        assert!(t.is_trusted(wb, "disk") && t.is_trusted(wb, "edit1"));
        for i in 2..=6 {
            t.pretrust(wb, &format!("edit{i}")).unwrap();
        }
        assert!(t.is_trusted(wb, "edit6"));
        assert!(
            !t.is_trusted(wb, "disk"),
            "oldest entries are dropped past the cap"
        );
        // An explicit Enable resets to exactly one digest.
        t.trust(wb, "only").unwrap();
        assert!(!t.is_trusted(wb, "edit6") && t.is_trusted(wb, "only"));
        // Other workbooks are untouched.
        t.pretrust(Path::new("/w/b.xlsx"), "x").unwrap();
        assert!(t.is_trusted(wb, "only"));
    }

    #[test]
    fn persists_and_fails_closed() {
        let dir = tempfile::tempdir().unwrap();
        let file = dir.path().join("sub").join("trust.json");
        let wb = Path::new("/tmp/a.xlsx");
        let mut t = TrustStore::load(file.clone());
        t.trust(wb, "d1").unwrap();
        assert!(TrustStore::load(file.clone()).is_trusted(wb, "d1"));

        std::fs::write(&file, b"{ not json").unwrap();
        assert!(!TrustStore::load(file).is_trusted(wb, "d1"));
    }
}
