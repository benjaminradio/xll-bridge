//! Diagnostics, `%APPDATA%\xll-bridge\addin.log`, written only when there is something to report.
//!
//! Excel reports add-in crashes with a generic dialog and no detail. Routine breadcrumbs
//! ([`note`]) are kept in a small in-memory ring and cost nothing on disk; when something goes
//! wrong ([`line`]: a panic, an access violation, a failed write) the recent breadcrumbs are
//! flushed to the file ahead of the problem, so the log reads as "what happened, then what broke".
//! A normal session therefore never creates the folder. Failures to log are ignored: logging
//! must never be the thing that breaks.

use std::collections::VecDeque;
use std::io::Write;
use std::path::PathBuf;
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};

const MAX_BYTES: u64 = 256 * 1024;
const MAX_NOTES: usize = 64;

static NOTES: Mutex<VecDeque<String>> = Mutex::new(VecDeque::new());

pub fn path() -> Option<PathBuf> {
    crate::trust::TrustStore::default_path().and_then(|p| p.parent().map(|d| d.join("addin.log")))
}

fn stamp(msg: &str) -> String {
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_or(0, |d| d.as_secs());
    format!("[{secs}] {msg}")
}

/// Remembers a breadcrumb in memory only.
pub fn note(msg: &str) {
    if let Ok(mut notes) = NOTES.lock() {
        if notes.len() == MAX_NOTES {
            notes.pop_front();
        }
        notes.push_back(stamp(msg));
    }
}

/// Breadcrumbs currently held (oldest first); used by tests.
pub fn pending_notes() -> Vec<String> {
    NOTES
        .lock()
        .map(|n| n.iter().cloned().collect())
        .unwrap_or_default()
}

/// Writes a problem to the log file, preceded by the breadcrumbs that led up to it.
pub fn line(msg: &str) {
    let Some(path) = path() else { return };
    let earlier: Vec<String> = NOTES
        .lock()
        .map(|mut n| n.drain(..).collect())
        .unwrap_or_default();
    write_lines(&path, &earlier, &stamp(msg));
}

fn write_lines(path: &std::path::Path, earlier: &[String], problem: &str) {
    if let Some(dir) = path.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    if std::fs::metadata(path).is_ok_and(|m| m.len() > MAX_BYTES) {
        let _ = std::fs::remove_file(path);
    }
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
        for e in earlier {
            let _ = writeln!(f, "  {e}");
        }
        let _ = writeln!(f, "{problem}");
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn breadcrumbs_stay_in_memory_and_are_bounded() {
        // Other tests may add notes, so assert on our own markers only.
        for i in 0..(MAX_NOTES + 10) {
            note(&format!("crumb-test-{i}"));
        }
        let held = pending_notes();
        assert!(held.len() <= MAX_NOTES);
        assert!(held.last().is_some_and(|l| l.ends_with("crumb-test-73")));
        assert!(
            !held.iter().any(|l| l.ends_with("crumb-test-0")),
            "oldest were dropped"
        );
    }

    #[test]
    fn a_problem_flushes_the_breadcrumbs_before_it_and_only_then_creates_the_folder() {
        let dir = tempfile::tempdir().unwrap();
        let file = dir.path().join("sub").join("addin.log");
        assert!(
            !file.parent().unwrap().exists(),
            "nothing is created until something is written"
        );
        write_lines(
            &file,
            &["[1] opened".into(), "[2] clicked".into()],
            "[3] panic: boom",
        );
        let text = std::fs::read_to_string(&file).unwrap();
        let order: Vec<usize> = ["opened", "clicked", "panic: boom"]
            .iter()
            .map(|s| text.find(s).unwrap())
            .collect();
        assert!(order.windows(2).all(|w| w[0] < w[1]), "{text}");
    }
}
