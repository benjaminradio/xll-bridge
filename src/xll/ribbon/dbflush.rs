//! Stores changed embedded databases in the open workbook (see [`crate::dbstore`]).

use windows::Win32::System::Com::IDispatch;

use super::{cells, live, set_status_bar, with_runtime, workbook_name_by_path};
use crate::dbstore::{self, Flush};

enum Outcome {
    Stored,
    /// The workbook is not open any more: nothing to store into.
    Gone,
    Failed(String),
}

fn store(app: &IDispatch, f: &Flush) -> Outcome {
    let Some(book) = workbook_name_by_path(&f.workbook) else {
        return Outcome::Gone;
    };
    let wb = match cells::open_book(app, &book) {
        Ok(wb) => wb,
        Err(e) => return Outcome::Failed(e.message),
    };
    let mut bundle = live::read(&wb).ok().flatten().unwrap_or_default();
    if let Err(e) = bundle.put_file(&f.path, f.bytes.clone(), true) {
        return Outcome::Failed(e.to_string());
    }
    if let Err(e) = live::write(&wb, &bundle) {
        return Outcome::Failed(e.to_string());
    }
    // Only a data file changed, so the running scripts and their trust are unaffected.
    match with_runtime(|rt| rt.apply_live(&f.workbook, &bundle)) {
        Ok(_) => Outcome::Stored,
        Err(e) => Outcome::Failed(e.to_string()),
    }
}

/// Called by the 200 ms timer.
pub(super) fn tick(app: &IDispatch) {
    for f in dbstore::take_flushes(dbstore::FLUSH_AFTER) {
        match store(app, &f) {
            Outcome::Stored => set_status_bar(
                app,
                &format!(
                    "Lua: database {} stored in the workbook (save it to keep)",
                    f.path
                ),
            ),
            Outcome::Gone => {}
            Outcome::Failed(e) => {
                f.failed();
                set_status_bar(
                    app,
                    &format!("Lua: could not store database {}: {e}", f.path),
                );
            }
        }
    }
}
