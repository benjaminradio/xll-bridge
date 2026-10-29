//! "Sync folder": keeps the active workbook's embedded Lua files in step with a folder on disk,
//! so scripts can be edited in an external editor. The folder is the source of truth; the
//! pure logic lives in [`crate::sync`].

use std::cell::RefCell;
use std::collections::HashMap;
use std::path::Path;

use windows::Win32::System::Com::IDispatch;
use windows::Win32::UI::WindowsAndMessaging::{
    IDNO, IDYES, MB_DEFBUTTON2, MB_DEFBUTTON3, MB_ICONERROR, MB_ICONINFORMATION, MB_ICONQUESTION,
    MB_OK, MB_SETFOREGROUND, MB_YESNO, MB_YESNOCANCEL, MessageBoxW,
};
use windows::core::PCWSTR;

use super::{
    active_workbook_path, cells, dialogs, dispatch, excel_hwnd, live, message, set_status_bar,
    with_runtime, workbook_name_by_path,
};
use crate::sync::{Links, Poll, Snapshot, Watcher};
use crate::trust::normalize_path;

/// The timer ticks every 200 ms; folders are looked at every this many ticks.
const POLL_EVERY: u32 = 5;

#[derive(Default)]
struct State {
    links: Option<Links>,
    watchers: HashMap<String, Watcher>,
    ticks: u32,
}

thread_local! {
    static STATE: RefCell<State> = RefCell::new(State::default());
}

pub(super) fn shutdown() {
    STATE.with(|s| *s.borrow_mut() = State::default());
}

fn name_of(p: &Path) -> String {
    p.file_name()
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_default()
}

/// Starts running `bundle` (already stored in the open workbook) as the workbook's scripts: the new
/// code is trusted (the user wrote it here), results and streams of the old code are dropped and the
/// workbook recalculates. A change that only touches data files leaves the running state alone.
pub(super) fn run_live(
    app: &IDispatch,
    workbook: &Path,
    bundle: &crate::vfs::Bundle,
) -> Result<(), String> {
    let code_changed =
        with_runtime(|rt| rt.apply_live(workbook, bundle)).map_err(|e| e.to_string())?;
    if code_changed {
        crate::jobs::global().forget_workbook(workbook);
        crate::streams::global().cancel_workbook(workbook);
        let _ = dispatch::call(app, "CalculateFull", &[]);
    }
    Ok(())
}

/// Writes the folder's files into the open workbook and starts running them.
fn apply(app: &IDispatch, workbook: &Path, snap: &Snapshot) -> Result<(), String> {
    let book = workbook_name_by_path(workbook)
        .ok_or_else(|| "the workbook is not open in Excel".to_owned())?;
    let wb = cells::open_book(app, &book).map_err(|e| e.message)?;
    live::write(&wb, &snap.bundle).map_err(|e| format!("could not update the workbook: {e}"))?;
    run_live(app, workbook, &snap.bundle)
}

fn report(app: &IDispatch, folder: &Path, snap: &Snapshot) {
    let n = snap.bundle.file_count();
    let mut text = format!(
        "Lua: synced {n} file{} from {}",
        if n == 1 { "" } else { "s" },
        folder.display()
    );
    if !snap.skipped.is_empty() {
        text.push_str(&format!(" ({} skipped)", snap.skipped.len()));
    }
    set_status_bar(app, &text);
}

/// Applies a change found by a watcher, unless it would empty the workbook.
fn deliver(app: &IDispatch, workbook: &Path, folder: &Path, snap: Snapshot) {
    if snap.bundle.file_count() == 0 {
        set_status_bar(
            app,
            &format!(
                "Lua: {} has no files; the workbook keeps its own",
                folder.display()
            ),
        );
        return;
    }
    match apply(app, workbook, &snap) {
        Ok(()) => report(app, folder, &snap),
        Err(e) => {
            crate::log::note(&format!("sync: {e}"));
            set_status_bar(app, &format!("Lua: sync failed: {e}"));
        }
    }
}

/// Called by the 200 ms timer.
pub(super) fn tick(app: &IDispatch) {
    let due = STATE.with(|s| {
        let mut s = s.borrow_mut();
        s.ticks += 1;
        s.ticks.is_multiple_of(POLL_EVERY)
    });
    if !due {
        return;
    }
    STATE.with(|s| {
        s.borrow_mut().links.get_or_insert_with(Links::load);
    });
    let open = super::open_workbook_paths();
    for path in open {
        let key = normalize_path(&path);
        let folder = STATE.with(|s| {
            s.borrow()
                .links
                .as_ref()
                .and_then(|l| l.get(&path).map(Path::to_path_buf))
        });
        let Some(folder) = folder else {
            STATE.with(|s| s.borrow_mut().watchers.remove(&key));
            continue;
        };
        // A workbook seen for the first time is compared with its folder straight away.
        let fresh = STATE.with(|s| !s.borrow().watchers.contains_key(&key));
        if fresh {
            let mut w = Watcher::new(folder.clone());
            match w.sync_now() {
                Ok(snap) => {
                    let embedded = workbook_name_by_path(&path)
                        .and_then(|n| cells::open_book(app, &n).ok())
                        .and_then(|wb| live::read(&wb).ok().flatten())
                        .map(|b| b.content_digest());
                    if embedded.as_deref() != Some(snap.bundle.content_digest().as_str()) {
                        deliver(app, &path, &folder, snap);
                    }
                }
                Err(e) => set_status_bar(app, &format!("Lua: sync paused: {e}")),
            }
            STATE.with(|s| s.borrow_mut().watchers.insert(key, w));
            continue;
        }
        let polled = STATE.with(|s| s.borrow_mut().watchers.get_mut(&key).map(Watcher::poll));
        match polled {
            Some(Poll::Changed(snap)) => deliver(app, &path, &folder, snap),
            Some(Poll::Failed(e)) => set_status_bar(app, &format!("Lua: sync paused: {e}")),
            _ => {}
        }
    }
}

enum Choice {
    Yes,
    No,
    Cancel,
}

fn ask3(owner: windows::Win32::Foundation::HWND, text: &str, caption: &str) -> Choice {
    let (t, c): (Vec<u16>, Vec<u16>) = (
        text.encode_utf16().chain(Some(0)).collect(),
        caption.encode_utf16().chain(Some(0)).collect(),
    );
    // SAFETY: both buffers are NUL-terminated and outlive the call.
    let r = unsafe {
        MessageBoxW(
            Some(owner),
            PCWSTR(t.as_ptr()),
            PCWSTR(c.as_ptr()),
            MB_YESNOCANCEL | MB_ICONQUESTION | MB_DEFBUTTON3 | MB_SETFOREGROUND,
        )
    };
    match r {
        x if x == IDYES => Choice::Yes,
        x if x == IDNO => Choice::No,
        _ => Choice::Cancel,
    }
}

/// The ribbon button: link a folder, or manage the existing link.
pub(super) fn on_click(app: &IDispatch) {
    let owner = excel_hwnd(app);
    let Some(path) = active_workbook_path(app) else {
        message(
            owner,
            "Save the workbook first; a folder is linked to a saved file.",
            "Lua",
            MB_OK | MB_ICONINFORMATION,
        );
        return;
    };
    let linked = STATE.with(|s| {
        s.borrow_mut()
            .links
            .get_or_insert_with(Links::load)
            .get(&path)
            .map(Path::to_path_buf)
    });
    match linked {
        Some(folder) => manage(app, owner, &path, &folder),
        None => link(app, owner, &path),
    }
}

fn manage(app: &IDispatch, owner: windows::Win32::Foundation::HWND, path: &Path, folder: &Path) {
    let text = format!(
        "\"{}\" follows the folder:\n\n{}\n\nYes: stop following it (the files stay in the workbook)\nNo: sync now\nCancel: do nothing",
        name_of(path),
        folder.display()
    );
    match ask3(owner, &text, "Sync folder") {
        Choice::Yes => {
            STATE.with(|s| {
                let mut s = s.borrow_mut();
                if let Some(l) = s.links.as_mut() {
                    l.remove(path);
                    l.save();
                }
                s.watchers.remove(&normalize_path(path));
            });
            set_status_bar(app, "Lua: folder unlinked");
        }
        Choice::No => {
            let mut w = Watcher::new(folder.to_path_buf());
            match w.sync_now() {
                Ok(snap) => deliver(app, path, folder, snap),
                Err(e) => set_status_bar(app, &format!("Lua: sync failed: {e}")),
            }
            STATE.with(|s| s.borrow_mut().watchers.insert(normalize_path(path), w));
        }
        Choice::Cancel => {}
    }
}

fn link(app: &IDispatch, owner: windows::Win32::Foundation::HWND, path: &Path) {
    let Some(folder) = dialogs::pick_folder(owner, "Choose the folder that holds your Lua files")
    else {
        return;
    };
    let mut watcher = Watcher::new(folder.clone());
    let snap = match watcher.sync_now() {
        Ok(s) => s,
        Err(e) => {
            message(owner, &e, "Sync folder", MB_OK | MB_ICONERROR);
            return;
        }
    };
    let embedded = workbook_name_by_path(path)
        .and_then(|n| cells::open_book(app, &n).ok())
        .and_then(|wb| live::read(&wb).ok().flatten());
    // An empty folder is the way to start from what the workbook already holds: offer to export
    // the embedded files into it, then follow it from there.
    let mut snap = snap;
    if snap.bundle.file_count() == 0 {
        let Some(held) = embedded.as_ref().filter(|b| b.file_count() > 0) else {
            message(
                owner,
                &format!("{} has no files to sync.", folder.display()),
                "Sync folder",
                MB_OK | MB_ICONINFORMATION,
            );
            return;
        };
        let text = format!(
            "{} is empty.\n\nExport the {} file(s) embedded in \"{}\" into it, and keep the workbook in step with that folder from now on?",
            folder.display(),
            held.file_count(),
            name_of(path)
        );
        if !message(
            owner,
            &text,
            "Sync folder",
            MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2,
        ) {
            return;
        }
        if let Err(e) = held.export_all(&folder) {
            message(
                owner,
                &format!("Could not export the files: {e}"),
                "Sync folder",
                MB_OK | MB_ICONERROR,
            );
            return;
        }
        match watcher.sync_now() {
            Ok(s) if s.bundle.file_count() > 0 => snap = s,
            _ => {
                message(
                    owner,
                    "The export did not produce any files.",
                    "Sync folder",
                    MB_OK | MB_ICONERROR,
                );
                return;
            }
        }
        deliver(app, path, &folder, snap);
        STATE.with(|s| {
            let mut s = s.borrow_mut();
            let l = s.links.get_or_insert_with(Links::load);
            l.set(path, folder);
            l.save();
            s.watchers.insert(normalize_path(path), watcher);
        });
        return;
    }
    let current = embedded.as_ref().map_or(0, |b| b.file_count());
    let text = format!(
        "Keep \"{}\" in step with this folder?\n\n{}\n\nIts {} file(s) will replace the {current} file(s) embedded in the workbook now, \
         and every change you save there is copied into the workbook within a couple of seconds. \
         Edits made in the Files panel will be overwritten. Lua will run these files without asking again.",
        name_of(path),
        folder.display(),
        snap.bundle.file_count(),
    );
    if !message(
        owner,
        &text,
        "Sync folder",
        MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2,
    ) {
        return;
    }
    deliver(app, path, &folder, snap);
    STATE.with(|s| {
        let mut s = s.borrow_mut();
        let l = s.links.get_or_insert_with(Links::load);
        l.set(path, folder);
        l.save();
        s.watchers.insert(normalize_path(path), watcher);
    });
}
