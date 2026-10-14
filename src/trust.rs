//! Trust store: Lua in a workbook never runs until the user enables it.
//!
//! Trust is pinned to (workbook path, script digest). Editing any script, or moving the file,
//! invalidates it and the banner comes back — the same shape as VBA's trusted documents.

use std::path::{Path, PathBuf};

use serde::{Deserialize, Serialize};

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
        TrustStore { path: None, entries: Vec::new() }
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
        TrustStore { path: Some(path), entries }
    }

    pub fn is_trusted(&self, workbook: &Path, digest: &str) -> bool {
        let key = normalize_path(workbook);
        self.entries.iter().any(|e| e.workbook == key && e.digest == digest)
    }

    /// Trusts this exact script set, replacing any earlier digest for the same workbook.
    pub fn trust(&mut self, workbook: &Path, digest: &str) -> std::io::Result<()> {
        let key = normalize_path(workbook);
        self.entries.retain(|e| e.workbook != key);
        self.entries.push(Entry { workbook: key, digest: digest.to_owned() });
        self.save()
    }

    fn save(&self) -> std::io::Result<()> {
        let Some(path) = &self.path else { return Ok(()) };
        if let Some(dir) = path.parent() {
            std::fs::create_dir_all(dir)?;
        }
        let json = serde_json::to_vec_pretty(&Disk { trusted: self.entries.clone() })
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
        assert!(!t.is_trusted(wb, "bb"), "edited scripts must not stay trusted");
        assert!(!t.is_trusted(Path::new(r"C:\Other\Book.xlsx"), "aa"));
        t.trust(wb, "bb").unwrap();
        assert!(!t.is_trusted(wb, "aa"), "new digest replaces the old one");
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
