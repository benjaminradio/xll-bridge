//! A tiny append-only diagnostic log, `%APPDATA%\xll-bridge\addin.log`.
//!
//! Excel reports add-in crashes with a generic dialog and no detail; these breadcrumbs (and the
//! panic/exception hooks in the Windows layer) are what make them diagnosable. Failures to log
//! are ignored: logging must never be the thing that breaks.

use std::io::Write;
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

const MAX_BYTES: u64 = 256 * 1024;

pub fn path() -> Option<PathBuf> {
    crate::trust::TrustStore::default_path().and_then(|p| p.parent().map(|d| d.join("addin.log")))
}

/// Appends `[unix-seconds] message`. The file is restarted when it grows past 256 KB.
pub fn line(msg: &str) {
    let Some(path) = path() else { return };
    if let Some(dir) = path.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    if std::fs::metadata(&path).is_ok_and(|m| m.len() > MAX_BYTES) {
        let _ = std::fs::remove_file(&path);
    }
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_or(0, |d| d.as_secs());
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
        let _ = writeln!(f, "[{secs}] {msg}");
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn writes_and_does_not_panic_without_a_home() {
        // Whatever the environment, logging is best-effort and infallible.
        super::line("test line");
    }
}
