//! The "Lua files" task pane: an Office custom task pane hosting the built-in browser control
//! (`Shell.Explorer.2`) that renders `panel.html`.
//!
//! Everything runs on Excel's UI thread from the ribbon's timer: no event sinks, no threads.
//! The page queues user commands, the timer pulls them with `pull()`, performs them (dialogs,
//! edits to the open workbook's custom XML part) and pushes fresh state with `setState(...)`.

use std::cell::RefCell;
use std::ffi::c_void;
use std::path::PathBuf;
use std::time::{Duration, Instant};

use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::{VARIANT, VT_ERROR};
use windows::Win32::UI::WindowsAndMessaging::{MB_DEFBUTTON2, MB_ICONWARNING, MB_OK, MB_YESNO};
use windows::core::{BSTR, Interface, Result};

use super::{ICTPFactory, dialogs, dispatch, live, message};
use crate::panel::{self, Command, PanelState};
use crate::ui::{self, View};
use crate::vfs::{Bundle, Kind};
use crate::xll::with_runtime;

const PROGID: &str = "Shell.Explorer.2";
const TITLE: &str = "Lua files";
const PANE_WIDTH: i32 = 340;
const DOCK_RIGHT: i32 = 2; // msoCTPDockPositionRight
const READY_STATE_COMPLETE: i32 = 4;
const LIVE_REFRESH: Duration = Duration::from_secs(4);
const MESSAGE_TTL: Duration = Duration::from_secs(7);

struct Pane {
    window: i32,
    ctp: IDispatch,
    browser: IDispatch,
    ready: bool,
    /// Last state text sent to this pane's page.
    pushed: String,
}

#[derive(Default)]
struct Panel {
    factory: Option<ICTPFactory>,
    panes: Vec<Pane>,
    busy: bool,
    /// Identity of the workbook the cached data belongs to.
    key: String,
    live: Option<Bundle>,
    live_read: Option<Instant>,
    message: Option<(bool, String, Instant)>,
    /// Digests of file sets authored in a workbook that had no path yet.
    authored_unsaved: Vec<String>,
}

thread_local! {
    static PANEL: RefCell<Panel> = RefCell::new(Panel::default());
}

pub fn set_factory(f: ICTPFactory) {
    PANEL.with(|p| p.borrow_mut().factory = Some(f));
}

pub fn shutdown() {
    PANEL.with(|p| *p.borrow_mut() = Panel::default());
}

fn missing() -> VARIANT {
    let mut v = VARIANT::default();
    // SAFETY: initialising a default VARIANT as VT_ERROR / DISP_E_PARAMNOTFOUND ("argument omitted").
    unsafe {
        let inner = &mut *v.Anonymous.Anonymous;
        inner.vt = VT_ERROR;
        inner.Anonymous.scode = 0x8002_0004u32 as i32;
    }
    v
}

fn flag(obj: &IDispatch, name: &str) -> Option<bool> {
    dispatch::get(obj, name)
        .ok()
        .and_then(|v| bool::try_from(&v).ok())
}

fn text(obj: &IDispatch, name: &str) -> Option<String> {
    dispatch::get(obj, name)
        .ok()
        .and_then(|v| BSTR::try_from(&v).ok())
        .map(|b| b.to_string())
}

fn active_window(app: &IDispatch) -> Option<(IDispatch, i32)> {
    let w = IDispatch::try_from(&dispatch::get(app, "ActiveWindow").ok()?).ok()?;
    let h = dispatch::get(&w, "Hwnd")
        .ok()
        .and_then(|v| i32::try_from(&v).ok())
        .unwrap_or(0);
    Some((w, h))
}

fn create_pane(app: &IDispatch, factory: &ICTPFactory) -> Result<Pane> {
    let (win, hwnd) = active_window(app).map_or((None, 0), |(w, h)| (Some(w), h));
    let progid = BSTR::from(PROGID);
    let title = BSTR::from(TITLE);
    let parent = win.map_or_else(missing, VARIANT::from);
    let mut out: *mut c_void = std::ptr::null_mut();
    // SAFETY: arguments are live BSTRs / a VARIANT; `out` receives an owned _CustomTaskPane.
    unsafe {
        factory
            .CreateCTP(progid.as_ptr(), title.as_ptr(), parent, &mut out)
            .ok()?
    };
    // SAFETY: on success `out` is an owned COM pointer.
    let ctp: IDispatch = unsafe { windows::core::IUnknown::from_raw(out) }.cast()?;
    dispatch::put(&ctp, "Visible", VARIANT::from(true))?;
    let _ = dispatch::put(&ctp, "DockPosition", VARIANT::from(DOCK_RIGHT));
    let _ = dispatch::put(&ctp, "Width", VARIANT::from(PANE_WIDTH));
    let browser = IDispatch::try_from(&dispatch::get(&ctp, "ContentControl")?)?;
    // No script-error dialogs inside Excel, whatever the page does.
    let _ = dispatch::put(&browser, "Silent", VARIANT::from(true));
    dispatch::call(&browser, "Navigate2", &[VARIANT::from("about:blank")])?;
    Ok(Pane {
        window: hwnd,
        ctp,
        browser,
        ready: false,
        pushed: String::new(),
    })
}

/// Shows the pane for the active window (creating it on first use), or hides it if it is open.
pub fn toggle(app: &IDispatch) {
    let owner = super::excel_hwnd(app);
    let result = PANEL.with(|p| -> std::result::Result<(), String> {
        let mut p = p.borrow_mut();
        let window = active_window(app).map_or(0, |(_, h)| h);
        if let Some(pane) = p.panes.iter().find(|x| x.window == window) {
            let open = flag(&pane.ctp, "Visible").unwrap_or(false);
            dispatch::put(&pane.ctp, "Visible", VARIANT::from(!open)).map_err(|e| e.to_string())?;
            return Ok(());
        }
        let factory = p
            .factory
            .clone()
            .ok_or("Excel did not offer a task pane factory to this add-in.")?;
        let pane = create_pane(app, &factory)
            .map_err(|e| format!("could not create the task pane ({e})"))?;
        p.panes.push(pane);
        p.key.clear(); // force a fresh read of the workbook
        Ok(())
    });
    if let Err(msg) = result {
        message(
            owner,
            &format!("The Lua files panel is unavailable: {msg}"),
            "Lua",
            MB_OK | MB_ICONWARNING,
        );
    }
    tick(app);
}

// ------------------------------------------------------------------ the per-tick work

struct Active {
    wb: IDispatch,
    name: String,
    key: String,
    path: Option<PathBuf>,
}

fn active_workbook(app: &IDispatch) -> Option<Active> {
    let wb = IDispatch::try_from(&dispatch::get(app, "ActiveWorkbook").ok()?).ok()?;
    let name = text(&wb, "Name")?;
    let full = text(&wb, "FullName").unwrap_or_else(|| name.clone());
    let path = ["LocalFullName", "FullName"]
        .into_iter()
        .filter_map(|p| text(&wb, p))
        .map(PathBuf::from)
        .find(|p| p.is_absolute() && p.is_file());
    Some(Active {
        wb,
        name,
        key: full,
        path,
    })
}

fn say(p: &mut Panel, is_err: bool, text: String) {
    p.message = Some((is_err, text, Instant::now()));
}

pub fn tick(app: &IDispatch) {
    let go = PANEL.with(|p| {
        let mut p = p.borrow_mut();
        if p.busy || p.panes.is_empty() {
            return false;
        }
        p.busy = true;
        true
    });
    if !go {
        return;
    }
    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| run_tick(app)));
    PANEL.with(|p| p.borrow_mut().busy = false);
    let _ = outcome;
}

fn run_tick(app: &IDispatch) {
    let window = active_window(app).map_or(0, |(_, h)| h);
    // Create a pane for a window we haven't seen if the panel is open elsewhere (SDI Excel).
    let missing_pane = PANEL.with(|p| {
        let p = p.borrow();
        let any_open = p
            .panes
            .iter()
            .any(|x| flag(&x.ctp, "Visible").unwrap_or(false));
        any_open && !p.panes.iter().any(|x| x.window == window)
    });
    if missing_pane
        && let Some(factory) = PANEL.with(|p| p.borrow().factory.clone())
        && let Ok(pane) = create_pane(app, &factory)
    {
        PANEL.with(|p| p.borrow_mut().panes.push(pane));
    }

    // Only the pane of the active window is serviced, and only while it is open.
    let Some(idx) = PANEL.with(|p| {
        let p = p.borrow();
        p.panes
            .iter()
            .position(|x| x.window == window && flag(&x.ctp, "Visible").unwrap_or(false))
    }) else {
        return;
    };
    let (browser, ready) = PANEL.with(|p| {
        let p = p.borrow();
        (p.panes[idx].browser.clone(), p.panes[idx].ready)
    });
    let page = match page_window(&browser) {
        Some(w) => w,
        None => {
            // Blank document not ready yet.
            if !ready && load_page(&browser).is_ok() {
                PANEL.with(|p| {
                    let mut p = p.borrow_mut();
                    p.panes[idx].ready = true;
                    p.panes[idx].pushed.clear();
                });
            }
            return;
        }
    };
    if !page_has_function(&page, "pull") {
        // The about:blank document is up but our page isn't written yet (or was navigated away).
        if load_page(&browser).is_ok() {
            PANEL.with(|p| {
                let mut p = p.borrow_mut();
                p.panes[idx].ready = true;
                p.panes[idx].pushed.clear();
            });
        }
        return;
    }

    let active = active_workbook(app);
    adopt_authored(active.as_ref());

    // 1) user commands
    while let Some(line) = dispatch::call(&page, "pull", &[])
        .ok()
        .and_then(|v| BSTR::try_from(&v).ok())
        .map(|b| b.to_string())
    {
        if line.is_empty() {
            break;
        }
        if let Some(cmd) = panel::parse_command(&line) {
            handle(app, active.as_ref(), cmd);
        }
    }

    // 2) state
    push_state(idx, &page, active.as_ref());
}

fn page_window(browser: &IDispatch) -> Option<IDispatch> {
    if i32::try_from(&dispatch::get(browser, "ReadyState").ok()?).ok()? != READY_STATE_COMPLETE {
        return None;
    }
    let doc = IDispatch::try_from(&dispatch::get(browser, "Document").ok()?).ok()?;
    IDispatch::try_from(&dispatch::get(&doc, "parentWindow").ok()?).ok()
}

fn page_has_function(page: &IDispatch, name: &str) -> bool {
    let wide: Vec<u16> = name.encode_utf16().chain(Some(0)).collect();
    let names = [windows::core::PCWSTR(wide.as_ptr())];
    let mut id = 0i32;
    // SAFETY: one name, one id slot.
    unsafe {
        page.GetIDsOfNames(
            &windows::core::GUID::zeroed(),
            names.as_ptr(),
            1,
            0x0400,
            &mut id,
        )
    }
    .is_ok()
}

fn load_page(browser: &IDispatch) -> Result<()> {
    if i32::try_from(&dispatch::get(browser, "ReadyState")?)? != READY_STATE_COMPLETE {
        return Err(windows::core::Error::from_hresult(
            windows::core::HRESULT(0x8000_000Au32 as i32), // E_PENDING
        ));
    }
    let doc = IDispatch::try_from(&dispatch::get(browser, "Document")?)?;
    dispatch::call(&doc, "open", &[])?;
    dispatch::call(&doc, "write", &[VARIANT::from(panel::HTML)])?;
    dispatch::call(&doc, "close", &[])?;
    Ok(())
}

/// A file set the user built in an unsaved workbook is theirs: once it is saved, trust it.
fn adopt_authored(active: Option<&Active>) {
    let Some(path) = active.and_then(|a| a.path.as_deref()) else {
        return;
    };
    let digests = PANEL.with(|p| p.borrow().authored_unsaved.clone());
    if digests.is_empty() {
        return;
    }
    with_runtime(|rt| {
        if let Some(d) = rt.disk_digest(path)
            && digests.contains(&d)
        {
            rt.adopt_authored(path, &d);
        }
    });
}

fn refresh_live(active: &Active, force: bool) -> Option<Bundle> {
    PANEL.with(|p| {
        let mut p = p.borrow_mut();
        let stale =
            force || p.key != active.key || p.live_read.is_none_or(|t| t.elapsed() > LIVE_REFRESH);
        if stale {
            p.live = live::read(&active.wb).ok().flatten();
            p.live_read = Some(Instant::now());
            p.key = active.key.clone();
        }
        p.live.clone()
    })
}

fn push_state(idx: usize, page: &IDispatch, active: Option<&Active>) {
    let (message, fresh) = PANEL.with(|p| {
        let mut p = p.borrow_mut();
        if p.message
            .as_ref()
            .is_some_and(|(_, _, t)| t.elapsed() > MESSAGE_TTL)
        {
            p.message = None;
        }
        let fresh = p.panes[idx].pushed.is_empty()
            || !p.panes[idx]
                .pushed
                .contains(&format!("wb={:?};", active.map(|a| &a.key)));
        (p.message.as_ref().map(|(e, t, _)| (*e, t.clone())), fresh)
    });

    let js = match active {
        None => {
            let view = ui::view_without_file();
            panel::state_js(&PanelState {
                workbook: None,
                ok: false,
                hint: "",
                view: &view,
                dirty: false,
                fresh,
                message,
                bundle: None,
            })
        }
        Some(a) => {
            let bundle = refresh_live(a, false);
            let (view, disk): (View, Option<String>) = match &a.path {
                Some(p) => with_runtime(|rt| (ui::view(&rt.status(p)), rt.disk_digest(p))),
                None => (ui::view_without_file(), None),
            };
            let live_digest = bundle.as_ref().map(Bundle::digest);
            let dirty = live_digest != disk;
            let empty = Bundle::new();
            panel::state_js(&PanelState {
                workbook: Some(&a.name),
                ok: true,
                hint: "",
                view: &view,
                dirty,
                fresh,
                message,
                bundle: Some(bundle.as_ref().unwrap_or(&empty)),
            })
        }
    };
    let signature =
        format!("wb={:?};{js}", active.map(|a| &a.key)).replace("fresh:true", "fresh:false");
    let changed = PANEL.with(|p| p.borrow().panes[idx].pushed != signature);
    if changed {
        let _ = dispatch::call(
            page,
            "execScript",
            &[VARIANT::from(js.as_str()), VARIANT::from("JavaScript")],
        );
        PANEL.with(|p| p.borrow_mut().panes[idx].pushed = signature);
    }
}

// ------------------------------------------------------------------ commands

enum Outcome {
    Nothing,
    /// A new tree to write into the workbook, with the status line to show.
    Edit(Bundle, String),
    Note(String),
    Error(String),
}

fn edit(current: &Bundle, cmd: &Command) -> Outcome {
    let mut b = current.clone();
    match panel::apply_edit(&mut b, cmd) {
        Ok(msg) => Outcome::Edit(b, msg),
        Err(e) => Outcome::Error(e.to_string()),
    }
}

fn handle(app: &IDispatch, active: Option<&Active>, cmd: Command) {
    let owner = super::excel_hwnd(app);
    let Some(active) = active else { return };
    let current = refresh_live(active, true).unwrap_or_default();
    let outcome = match &cmd {
        Command::Refresh => Outcome::Nothing,
        Command::NewFolder { .. } | Command::Rename { .. } => edit(&current, &cmd),
        Command::Delete { path } => {
            let what = panel::delete_summary(&current, path);
            let sure = message(
                owner,
                &format!("Delete {what} from this workbook?"),
                "Delete",
                MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2,
            );
            if sure {
                edit(&current, &cmd)
            } else {
                Outcome::Nothing
            }
        }
        Command::AddFiles { dest } => add(&current, dest, dialogs::pick_files(owner)),
        Command::AddFolder { dest } => add(&current, dest, dialogs::pick_folders(owner)),
        Command::Export { path } => export(owner, &current, path),
    };
    match outcome {
        Outcome::Nothing => {}
        Outcome::Edit(bundle, msg) => commit(active, bundle, msg),
        Outcome::Note(n) => PANEL.with(|p| say(&mut p.borrow_mut(), false, n)),
        Outcome::Error(e) => PANEL.with(|p| say(&mut p.borrow_mut(), true, e)),
    }
}

fn add(current: &Bundle, dest: &str, picked: Vec<PathBuf>) -> Outcome {
    if picked.is_empty() {
        return Outcome::Nothing;
    }
    let mut b = current.clone();
    let (mut files, mut replaced, mut skipped) = (0, 0, Vec::new());
    for p in &picked {
        match b.import_path(p, dest, true) {
            Ok(r) => {
                files += r.files;
                replaced += r.replaced;
                skipped.extend(r.skipped);
            }
            Err(e) => return Outcome::Error(e.to_string()),
        }
    }
    let mut msg = format!("Added {files} file{}", if files == 1 { "" } else { "s" });
    if replaced > 0 {
        msg.push_str(&format!(" ({replaced} replaced)"));
    }
    if !skipped.is_empty() {
        msg.push_str(&format!(
            " \u{2014} skipped {}: {}",
            skipped.len(),
            skipped
                .iter()
                .take(2)
                .cloned()
                .collect::<Vec<_>>()
                .join("; ")
        ));
    }
    Outcome::Edit(b, msg)
}

fn export(owner: HWND, b: &Bundle, path: &str) -> Outcome {
    let done = |r: std::result::Result<String, crate::vfs::VfsError>| match r {
        Ok(m) => Outcome::Note(m),
        Err(e) => Outcome::Error(e.to_string()),
    };
    match b.kind_of(path) {
        Some(Kind::File) => {
            let name = path.rsplit('/').next().unwrap_or(path);
            match dialogs::save_file(owner, name) {
                None => Outcome::Nothing,
                Some(dest) => done(
                    b.export_file_as(path, &dest)
                        .map(|_| format!("Exported to {}", dest.display())),
                ),
            }
        }
        Some(Kind::Dir) => match dialogs::pick_folder(owner, "Choose where to export the folder") {
            None => Outcome::Nothing,
            Some(dest) => done(
                b.export_to(path, &dest)
                    .map(|p| format!("Exported to {}", p.display())),
            ),
        },
        None => Outcome::Error(format!("'{path}' no longer exists")),
    }
}

/// Writes the new tree into the workbook and updates trust and the cached copy.
fn commit(active: &Active, bundle: Bundle, mut msg: String) {
    if let Err(e) = live::write(&active.wb, &bundle) {
        PANEL.with(|p| {
            say(
                &mut p.borrow_mut(),
                true,
                format!("Could not update the workbook: {e}"),
            )
        });
        return;
    }
    let digest = bundle.digest();
    if bundle.file_count() > 0 {
        match &active.path {
            Some(path) => {
                if !with_runtime(|rt| rt.note_edit(path, &digest)) {
                    msg.push_str(". Lua stays disabled until you enable it with the ribbon button");
                }
            }
            None => PANEL.with(|p| {
                let mut p = p.borrow_mut();
                p.authored_unsaved.push(digest);
                let n = p.authored_unsaved.len();
                if n > 8 {
                    p.authored_unsaved.drain(..n - 8);
                }
            }),
        }
    }
    PANEL.with(|p| {
        let mut p = p.borrow_mut();
        p.live = (!bundle.is_empty()).then_some(bundle);
        p.live_read = Some(Instant::now());
        say(&mut p, false, msg);
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::testutil::bundle_of;
    use windows::Win32::System::Com::{
        CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx,
    };
    use windows::core::w;

    /// A standalone MSHTML document (what the task pane's browser control renders with).
    /// `None` when the machine has no MSHTML, in which case the page tests are skipped.
    fn html_document() -> Option<IDispatch> {
        // SAFETY: standard COM bootstrap on the test thread.
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            let clsid = windows::Win32::System::Com::CLSIDFromProgID(w!("htmlfile")).ok()?;
            CoCreateInstance::<_, IDispatch>(&clsid, None, CLSCTX_INPROC_SERVER).ok()
        }
    }

    fn count(doc: &IDispatch, id: &str) -> i32 {
        let el = IDispatch::try_from(
            &dispatch::call(doc, "getElementById", &[VARIANT::from(id)]).unwrap(),
        )
        .unwrap();
        let kids = IDispatch::try_from(&dispatch::get(&el, "childNodes").unwrap()).unwrap();
        i32::try_from(&dispatch::get(&kids, "length").unwrap()).unwrap()
    }

    fn run(win: &IDispatch, js: &str) {
        dispatch::call(
            win,
            "execScript",
            &[VARIANT::from(js), VARIANT::from("JavaScript")],
        )
        .unwrap_or_else(|e| panic!("script failed ({e}): {js}"));
    }

    #[test]
    fn the_page_renders_state_and_queues_commands() {
        let Some(doc) = html_document() else {
            eprintln!("MSHTML not available on this machine: skipping page test");
            return;
        };
        dispatch::call(&doc, "open", &[]).unwrap();
        dispatch::call(&doc, "write", &[VARIANT::from(panel::HTML)]).unwrap();
        dispatch::call(&doc, "close", &[]).unwrap();
        let win = IDispatch::try_from(&dispatch::get(&doc, "parentWindow").unwrap()).unwrap();
        assert!(page_has_function(&win, "pull") && page_has_function(&win, "setState"));

        let bundle = bundle_of(&[
            ("main.lua", "x"),
            ("lib/util.lua", "y"),
            ("lib/deep/z.lua", "z"),
        ]);
        let view = ui::view(&crate::runtime::Status::Blocked {
            names: vec!["main.lua".into()],
        });
        let js = panel::state_js(&PanelState {
            workbook: Some("Book1.xlsx"),
            ok: true,
            hint: "",
            view: &view,
            dirty: true,
            fresh: true,
            message: None,
            bundle: Some(&bundle),
        });
        run(&win, &js);
        // Collapsed: the `lib` folder and `main.lua` at the top level.
        assert_eq!(count(&doc, "tree"), 2);
        // Expanding shows the folder's children.
        run(&win, "expanded['lib'] = true; render();");
        assert_eq!(count(&doc, "tree"), 4); // lib, lib/deep, lib/util.lua, main.lua

        // Buttons queue commands for the host, with the selected folder as the target.
        run(
            &win,
            "sel = { p: 'lib', k: 'd' }; send('addfiles', targetDir());",
        );
        run(&win, "send('rename', 'lib/util.lua', 'a b.lua');");
        let first = BSTR::try_from(&dispatch::call(&win, "pull", &[]).unwrap())
            .unwrap()
            .to_string();
        let second = BSTR::try_from(&dispatch::call(&win, "pull", &[]).unwrap())
            .unwrap()
            .to_string();
        let empty = BSTR::try_from(&dispatch::call(&win, "pull", &[]).unwrap())
            .unwrap()
            .to_string();
        assert_eq!(
            panel::parse_command(&first),
            Some(Command::AddFiles { dest: "lib".into() })
        );
        assert_eq!(
            panel::parse_command(&second),
            Some(Command::Rename {
                path: "lib/util.lua".into(),
                name: "a b.lua".into()
            })
        );
        assert_eq!(empty, "");

        // Inline editing renders an input, and hostile names are text, never markup.
        run(&win, "startEdit('new', '');");
        let tree = IDispatch::try_from(
            &dispatch::call(&doc, "getElementById", &[VARIANT::from("tree")]).unwrap(),
        )
        .unwrap();
        let html = BSTR::try_from(&dispatch::get(&tree, "innerHTML").unwrap())
            .unwrap()
            .to_string();
        assert!(html.to_lowercase().contains("<input"), "{html}");

        // Names with markup-significant characters render as plain text.
        let odd = bundle_of(&[("a&b'c.lua", "x")]);
        let js = panel::state_js(&PanelState {
            workbook: Some("Book1.xlsx"),
            ok: true,
            hint: "",
            view: &view,
            dirty: false,
            fresh: true,
            message: None,
            bundle: Some(&odd),
        });
        run(&win, &js);
        let text = BSTR::try_from(&dispatch::get(&tree, "innerText").unwrap())
            .unwrap()
            .to_string();
        assert!(text.contains("a&b'c.lua"), "{text}");
    }
}
