//! The "Lua files" panel: a native tool window (owned by the Excel main window, so it floats
//! above Excel and minimises with it) holding a tree view and a few buttons.
//!
//! No ActiveX, no web control and no Office task-pane factory are involved — just a window and
//! the standard tree control — because those were the part of the earlier design that could take
//! Excel down. Everything runs on Excel's UI thread: user actions become [`Command`]s on a queue
//! that is drained from a posted message (never from inside a control notification), and the
//! ribbon's timer calls [`tick`] to follow the active workbook.
//!
//! Rule for this file: never hold a `RefCell` borrow across a call that can send messages (the
//! window procedure re-enters), and keep window handles in `Cell`s.

use std::cell::{Cell, RefCell};
use std::collections::{HashMap, HashSet, VecDeque};
use std::ffi::c_void;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::PathBuf;
use std::time::{Duration, Instant};

use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, COLOR_3DFACE, CreateFontIndirectW, CreateSolidBrush, DEFAULT_GUI_FONT,
    DT_END_ELLIPSIS, DT_LEFT, DT_NOPREFIX, DT_SINGLELINE, DT_VCENTER, DeleteObject, DrawTextW,
    Ellipse, EndPaint, FillRect, GetObjectW, GetStockObject, GetSysColor, GetSysColorBrush, HDC,
    HFONT, HGDIOBJ, InvalidateRect, LOGFONTW, NULL_PEN, PAINTSTRUCT, SelectObject, SetBkMode,
    SetTextColor, TRANSPARENT,
};
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{
    HTREEITEM, ICC_TREEVIEW_CLASSES, INITCOMMONCONTROLSEX, InitCommonControlsEx, NM_DBLCLK, NMHDR,
    NMTREEVIEWW, NMTVDISPINFOW, NMTVKEYDOWN, TVE_COLLAPSE, TVE_EXPAND, TVGN_CARET, TVI_LAST,
    TVI_ROOT, TVIF_PARAM, TVIF_TEXT, TVINSERTSTRUCTW, TVINSERTSTRUCTW_0, TVITEMW, TVM_DELETEITEM,
    TVM_EDITLABELW, TVM_ENSUREVISIBLE, TVM_EXPAND, TVM_INSERTITEMW, TVM_SELECTITEM,
    TVN_BEGINLABELEDITW, TVN_ENDLABELEDITW, TVN_ITEMEXPANDEDW, TVN_KEYDOWN, TVN_SELCHANGEDW,
    TVS_EDITLABELS, TVS_HASBUTTONS, TVS_HASLINES, TVS_LINESATROOT, TVS_SHOWSELALWAYS, WC_TREEVIEW,
};
use windows::Win32::UI::HiDpi::GetDpiForWindow;
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, VK_DELETE, VK_F2};
use windows::Win32::UI::Shell::{DragAcceptFiles, DragFinish, DragQueryFileW, HDROP};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{PCWSTR, PWSTR, w};

use super::{dialogs, dispatch, live, message};
use crate::panel::{self, Command};
use crate::runtime::Status;
use crate::ui::{self, Light, View};
use crate::vfs::{Bundle, Entry, Kind};
use crate::xll::with_runtime;

const CLASS: PCWSTR = w!("XllBridgeFilesPanel");
const TITLE: PCWSTR = w!("Lua files");
const PANEL_WIDTH: i32 = 340;
const LIVE_REFRESH: Duration = Duration::from_secs(4);
const MESSAGE_TTL: Duration = Duration::from_secs(7);

const ID_TREE: usize = 200;
const ID_NEW: usize = 101;
const ID_ADD_FILES: usize = 102;
const ID_ADD_FOLDER: usize = 103;
const ID_EXPORT: usize = 104;
const ID_RENAME: usize = 105;
const ID_DELETE: usize = 106;
const BUTTONS: [(usize, PCWSTR); 6] = [
    (ID_NEW, w!("New folder")),
    (ID_ADD_FILES, w!("Add files\u{2026}")),
    (ID_ADD_FOLDER, w!("Add folder\u{2026}")),
    (ID_EXPORT, w!("Export\u{2026}")),
    (ID_RENAME, w!("Rename")),
    (ID_DELETE, w!("Delete")),
];

/// Posted to the panel to drain the command queue outside of any control notification.
const WM_QUEUE: u32 = WM_APP + 1;

#[derive(Clone, Copy)]
struct Handles {
    win: HWND,
    tree: HWND,
    foot: HWND,
}

struct Item {
    path: String,
    kind: Kind,
    handle: isize,
}

struct PendingNew {
    parent: String,
    handle: isize,
}

#[derive(Default)]
struct Model {
    items: Vec<Item>,
    expanded: HashSet<String>,
    selected: Option<String>,
    pending_new: Option<PendingNew>,
    // header / footer
    wb_name: String,
    view: Option<View>,
    foot: String,
    foot_is_error: bool,
    can_edit: bool,
    // controller
    key: String,
    live: Option<Bundle>,
    live_read: Option<Instant>,
    message: Option<(bool, String, Instant)>,
    signature: String,
    authored_unsaved: Vec<String>,
    busy: bool,
}

thread_local! {
    static HANDLES: Cell<Option<Handles>> = const { Cell::new(None) };
    static MODEL: RefCell<Model> = RefCell::new(Model::default());
    static QUEUE: RefCell<VecDeque<Command>> = const { RefCell::new(VecDeque::new()) };
    /// True while we rebuild the tree ourselves: its notifications must be ignored.
    static REBUILDING: Cell<bool> = const { Cell::new(false) };
    static EDITING: Cell<bool> = const { Cell::new(false) };
    static FONTS: Cell<Option<(isize, isize)>> = const { Cell::new(None) }; // (normal, bold)
}

fn handles() -> Option<Handles> {
    HANDLES.with(Cell::get)
}

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(Some(0)).collect()
}

fn from_wide_ptr(p: PWSTR) -> Option<String> {
    (!p.is_null()).then(|| unsafe { p.to_string().unwrap_or_default() })
}

pub fn shutdown() {
    if let Some(h) = handles() {
        // SAFETY: our own window, on the thread that created it.
        let _ = unsafe { DestroyWindow(h.win) };
    }
    HANDLES.with(|c| c.set(None));
    MODEL.with(|m| *m.borrow_mut() = Model::default());
    QUEUE.with(|q| q.borrow_mut().clear());
}

// ------------------------------------------------------------------ window creation

fn dpi_scale(hwnd: HWND, v: i32) -> i32 {
    // SAFETY: plain query; 0 for an invalid window, handled by the floor.
    let dpi = unsafe { GetDpiForWindow(hwnd) }.max(96) as i32;
    v * dpi / 96
}

fn make_fonts() -> (HFONT, HFONT) {
    if let Some((n, b)) = FONTS.with(Cell::get) {
        return (HFONT(n as *mut c_void), HFONT(b as *mut c_void));
    }
    // SAFETY: GDI font creation from the stock GUI font.
    unsafe {
        let stock = GetStockObject(DEFAULT_GUI_FONT);
        let mut lf = LOGFONTW::default();
        GetObjectW(
            stock,
            std::mem::size_of::<LOGFONTW>() as i32,
            Some((&mut lf as *mut LOGFONTW).cast()),
        );
        let normal = HFONT(stock.0);
        lf.lfWeight = 700;
        let bold = CreateFontIndirectW(&lf);
        FONTS.with(|f| f.set(Some((normal.0 as isize, bold.0 as isize))));
        (normal, bold)
    }
}

fn register_class() {
    thread_local!(static DONE: Cell<bool> = const { Cell::new(false) });
    if DONE.with(Cell::get) {
        return;
    }
    // SAFETY: standard window-class registration.
    unsafe {
        let hinst = GetModuleHandleW(None).unwrap_or_default();
        let wc = WNDCLASSW {
            lpfnWndProc: Some(wndproc),
            hInstance: hinst.into(),
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hbrBackground: GetSysColorBrush(COLOR_3DFACE),
            lpszClassName: CLASS,
            ..Default::default()
        };
        RegisterClassW(&wc);
    }
    DONE.with(|d| d.set(true));
}

/// Creates the panel window (hidden) on first use. `owner` is Excel's main window.
fn ensure_window(owner: HWND) -> Option<Handles> {
    if let Some(h) = handles() {
        return Some(h);
    }
    register_class();
    // SAFETY: Win32 window creation on the UI thread; all strings are static or outlive the calls.
    unsafe {
        let _ = InitCommonControlsEx(&INITCOMMONCONTROLSEX {
            dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_TREEVIEW_CLASSES,
        });
        let hinst = GetModuleHandleW(None).ok()?;
        // Start flush against the right edge of Excel, below the ribbon.
        let mut rc = RECT::default();
        let have_owner = GetWindowRect(owner, &mut rc).is_ok();
        let w_px = dpi_scale(owner, PANEL_WIDTH);
        let (x, y, h) = if have_owner {
            let y = rc.top + dpi_scale(owner, 140);
            (
                rc.right - w_px - dpi_scale(owner, 14),
                y,
                (rc.bottom - y - dpi_scale(owner, 60)).max(dpi_scale(owner, 360)),
            )
        } else {
            (100, 100, dpi_scale(owner, 520))
        };
        let win = CreateWindowExW(
            WS_EX_TOOLWINDOW,
            CLASS,
            TITLE,
            WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_CLIPCHILDREN,
            x,
            y,
            w_px,
            h,
            have_owner.then_some(owner),
            None,
            Some(hinst.into()),
            None,
        )
        .ok()?;
        let child = |ex, class: PCWSTR, text: PCWSTR, style: WINDOW_STYLE, id: usize| {
            CreateWindowExW(
                ex,
                class,
                text,
                WS_CHILD | WS_VISIBLE | style,
                0,
                0,
                10,
                10,
                Some(win),
                Some(HMENU(id as *mut c_void)),
                Some(hinst.into()),
                None,
            )
            .unwrap_or_default()
        };
        let tree_style = WINDOW_STYLE(
            TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS | TVS_SHOWSELALWAYS | TVS_EDITLABELS,
        ) | WS_TABSTOP;
        let tree = child(WS_EX_CLIENTEDGE, WC_TREEVIEW, w!(""), tree_style, ID_TREE);
        let foot = child(
            WINDOW_EX_STYLE(0),
            w!("STATIC"),
            w!(""),
            WINDOW_STYLE(0),
            300,
        );
        let (normal, _) = make_fonts();
        for (id, label) in BUTTONS {
            let b = child(
                WINDOW_EX_STYLE(0),
                w!("BUTTON"),
                label,
                WINDOW_STYLE(BS_PUSHBUTTON as u32) | WS_TABSTOP,
                id,
            );
            SendMessageW(
                b,
                WM_SETFONT,
                Some(WPARAM(normal.0 as usize)),
                Some(LPARAM(1)),
            );
        }
        SendMessageW(
            foot,
            WM_SETFONT,
            Some(WPARAM(normal.0 as usize)),
            Some(LPARAM(1)),
        );
        DragAcceptFiles(win, true);
        let h = Handles { win, tree, foot };
        HANDLES.with(|c| c.set(Some(h)));
        layout(&h);
        Some(h)
    }
}

fn layout(h: &Handles) {
    // SAFETY: our own windows.
    unsafe {
        let mut rc = RECT::default();
        if GetClientRect(h.win, &mut rc).is_err() {
            return;
        }
        let s = |v: i32| dpi_scale(h.win, v);
        let (w, total_h) = (rc.right, rc.bottom);
        let (header, pad, btn_h, foot_h) = (s(46), s(6), s(26), s(24));
        let bw = ((w - 4 * pad) / 3).max(40);
        for (i, (id, _)) in BUTTONS.iter().enumerate() {
            let (col, row) = ((i % 3) as i32, (i / 3) as i32);
            if let Ok(b) = GetDlgItem(Some(h.win), *id as i32) {
                let _ = MoveWindow(
                    b,
                    pad + col * (bw + pad),
                    header + pad + row * (btn_h + pad / 2),
                    bw,
                    btn_h,
                    true,
                );
            }
        }
        let tree_top = header + pad * 2 + btn_h * 2 + pad / 2;
        let _ = MoveWindow(
            h.tree,
            0,
            tree_top,
            w,
            (total_h - tree_top - foot_h).max(10),
            true,
        );
        let _ = MoveWindow(
            h.foot,
            pad,
            total_h - foot_h + s(4),
            (w - 2 * pad).max(10),
            foot_h - s(4),
            true,
        );
    }
}

// ------------------------------------------------------------------ show / hide

/// Shows the panel (creating it on first use), or hides it if it is open.
pub fn toggle(app: &IDispatch) {
    let owner = super::excel_hwnd(app);
    let Some(h) = ensure_window(owner) else {
        crate::log::line("panel: could not create the window");
        message(
            owner,
            "The Lua files panel could not be created.",
            "Lua",
            MB_OK | MB_ICONWARNING,
        );
        return;
    };
    // SAFETY: our own window.
    unsafe {
        if IsWindowVisible(h.win).as_bool() {
            let _ = ShowWindow(h.win, SW_HIDE);
        } else {
            let _ = ShowWindow(h.win, SW_SHOW);
            MODEL.with(|m| m.borrow_mut().signature.clear());
            tick(app);
        }
    }
}

// ------------------------------------------------------------------ the per-tick work

struct Active {
    wb: IDispatch,
    name: String,
    key: String,
    path: Option<PathBuf>,
}

fn text(obj: &IDispatch, name: &str) -> Option<String> {
    use windows::core::BSTR;
    dispatch::get(obj, name)
        .ok()
        .and_then(|v| BSTR::try_from(&v).ok())
        .map(|b| b.to_string())
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

fn say(is_err: bool, msg: String) {
    MODEL.with(|m| m.borrow_mut().message = Some((is_err, msg, Instant::now())));
}

/// Follows the active workbook and keeps the window current. Called from the ribbon's timer.
pub fn tick(app: &IDispatch) {
    let Some(h) = handles() else { return };
    // SAFETY: our own window.
    if !unsafe { IsWindowVisible(h.win) }.as_bool() || EDITING.with(Cell::get) {
        return;
    }
    let go = MODEL.with(|m| {
        let mut m = m.borrow_mut();
        !std::mem::replace(&mut m.busy, true)
    });
    if !go {
        return;
    }
    let outcome = catch_unwind(AssertUnwindSafe(|| refresh(app, h)));
    MODEL.with(|m| m.borrow_mut().busy = false);
    if outcome.is_err() {
        crate::log::line("panel: panic while refreshing (swallowed)");
    }
}

fn refresh(app: &IDispatch, h: Handles) {
    let active = active_workbook(app);
    adopt_authored(active.as_ref());

    let (view, bundle, dirty, name, key, can_edit) = match &active {
        None => (
            ui::view_without_file(),
            None,
            false,
            String::new(),
            String::new(),
            false,
        ),
        Some(a) => {
            let bundle = refresh_live(a, false);
            let (view, disk) = match &a.path {
                Some(p) => with_runtime(|rt| (ui::view(&rt.status(p)), rt.disk_digest(p))),
                None => (ui::view_without_file(), None),
            };
            let dirty = bundle.as_ref().map(Bundle::digest) != disk;
            (view, bundle, dirty, a.name.clone(), a.key.clone(), true)
        }
    };
    let message = MODEL.with(|m| {
        let mut m = m.borrow_mut();
        if m.message
            .as_ref()
            .is_some_and(|(_, _, t)| t.elapsed() > MESSAGE_TTL)
        {
            m.message = None;
        }
        m.message.as_ref().map(|(e, t, _)| (*e, t.clone()))
    });
    let digest = bundle.as_ref().map(Bundle::digest).unwrap_or_default();
    let signature = format!(
        "{key}|{digest}|{:?}|{}|{dirty}|{message:?}|{can_edit}",
        view.light, view.label
    );
    if MODEL.with(|m| m.borrow().signature == signature) {
        return;
    }
    let (foot, foot_is_error) = match (&message, dirty, can_edit) {
        (Some((e, t)), _, _) => (t.clone(), *e),
        (None, _, false) => (
            "Open a workbook to manage its embedded Lua files.".to_owned(),
            false,
        ),
        (None, true, _) => (
            "Unsaved changes \u{2014} save the workbook to apply them.".to_owned(),
            false,
        ),
        (None, false, _) => {
            let n = bundle.as_ref().map_or(0, Bundle::file_count);
            (format!("{n} file{}", if n == 1 { "" } else { "s" }), false)
        }
    };
    let entries = bundle.as_ref().map(Bundle::entries).unwrap_or_default();
    let roots: HashSet<String> = bundle
        .as_ref()
        .map(|b| b.root_scripts().iter().map(|f| f.path.clone()).collect())
        .unwrap_or_default();
    MODEL.with(|m| {
        let mut m = m.borrow_mut();
        if m.key != key {
            m.expanded.clear();
            m.selected = None;
        }
        m.wb_name = name;
        m.view = Some(view);
        m.foot = foot;
        m.foot_is_error = foot_is_error;
        m.can_edit = can_edit;
        m.signature = signature;
    });
    rebuild(h, &entries, &roots);
    update_chrome(h);
}

// ------------------------------------------------------------------ painting the window

fn update_chrome(h: Handles) {
    let (foot, can_edit, selected) = MODEL.with(|m| {
        let m = m.borrow();
        (m.foot.clone(), m.can_edit, m.selected.clone())
    });
    // SAFETY: our own windows.
    unsafe {
        let t = wide(&foot);
        let _ = SetWindowTextW(h.foot, PCWSTR(t.as_ptr()));
        for (id, _) in BUTTONS {
            let needs_selection = matches!(id, ID_EXPORT | ID_RENAME | ID_DELETE);
            let on = can_edit && (!needs_selection || selected.is_some());
            if let Ok(b) = GetDlgItem(Some(h.win), id as i32) {
                let _ = EnableWindow(b, on);
            }
        }
        let _ = InvalidateRect(Some(h.win), None, true);
    }
}

fn light_colour(l: Light) -> COLORREF {
    let (r, g, b) = match l {
        Light::Green => (46u32, 160u32, 67u32),
        Light::Yellow => (245, 190, 0),
        Light::Red => (214, 48, 49),
    };
    COLORREF(r | (g << 8) | (b << 16))
}

fn paint(hwnd: HWND) {
    let (name, view) = MODEL.with(|m| {
        let m = m.borrow();
        (m.wb_name.clone(), m.view.clone())
    });
    // SAFETY: standard WM_PAINT handling for our own window.
    unsafe {
        let mut ps = PAINTSTRUCT::default();
        let hdc: HDC = BeginPaint(hwnd, &mut ps);
        let s = |v: i32| dpi_scale(hwnd, v);
        let mut rc = RECT::default();
        let _ = GetClientRect(hwnd, &mut rc);
        let band = RECT {
            left: 0,
            top: 0,
            right: rc.right,
            bottom: s(46),
        };
        let bg = CreateSolidBrush(COLORREF(GetSysColor(COLOR_3DFACE)));
        FillRect(hdc, &band, bg);
        let _ = DeleteObject(HGDIOBJ(bg.0));

        let light = view.as_ref().map_or(Light::Green, |v| v.light);
        let dot = CreateSolidBrush(light_colour(light));
        let old_brush = SelectObject(hdc, HGDIOBJ(dot.0));
        let old_pen = SelectObject(hdc, GetStockObject(NULL_PEN));
        let d = s(14);
        let _ = Ellipse(hdc, s(10), s(16), s(10) + d, s(16) + d);
        SelectObject(hdc, old_brush);
        SelectObject(hdc, old_pen);
        let _ = DeleteObject(HGDIOBJ(dot.0));

        SetBkMode(hdc, TRANSPARENT);
        let (normal, bold) = make_fonts();
        let line = |font: HFONT, color: COLORREF, top: i32, s_: &str| {
            let mut buf: Vec<u16> = s_.encode_utf16().collect();
            let mut r = RECT {
                left: s(32),
                top,
                right: rc.right - s(8),
                bottom: top + s(18),
            };
            SelectObject(hdc, HGDIOBJ(font.0));
            SetTextColor(hdc, color);
            DrawTextW(
                hdc,
                &mut buf,
                &mut r,
                DT_LEFT | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX,
            );
        };
        let title = if name.is_empty() {
            "No workbook"
        } else {
            name.as_str()
        };
        line(bold, COLORREF(0x001f1f1f), s(5), title);
        line(
            normal,
            COLORREF(0x00555555),
            s(24),
            view.as_ref().map_or("", |v| v.screentip.as_str()),
        );
        let _ = EndPaint(hwnd, &ps);
    }
}

// ------------------------------------------------------------------ the tree

fn label(e: &Entry, roots: &HashSet<String>) -> String {
    let name = e.path.rsplit('/').next().unwrap_or(&e.path);
    match e.kind {
        Kind::Dir => name.to_owned(),
        Kind::File => {
            let size = if e.size < 1024 {
                format!("{} B", e.size)
            } else if e.size < 1 << 20 {
                format!("{:.1} KB", e.size as f64 / 1024.0)
            } else {
                format!("{:.1} MB", e.size as f64 / (1 << 20) as f64)
            };
            let run = if roots.contains(&e.path) {
                "  \u{25B6} runs at load"
            } else {
                ""
            };
            format!("{name}   ({size}){run}")
        }
    }
}

fn insert_item(tree: HWND, parent: isize, label: &str, param: isize) -> isize {
    let mut buf = wide(label);
    let tvi = TVINSERTSTRUCTW {
        hParent: HTREEITEM(parent),
        hInsertAfter: TVI_LAST,
        Anonymous: TVINSERTSTRUCTW_0 {
            item: TVITEMW {
                mask: TVIF_TEXT | TVIF_PARAM,
                pszText: PWSTR(buf.as_mut_ptr()),
                lParam: LPARAM(param),
                ..Default::default()
            },
        },
    };
    // SAFETY: `tvi` and `buf` outlive the synchronous message.
    unsafe {
        SendMessageW(
            tree,
            TVM_INSERTITEMW,
            None,
            Some(LPARAM(&tvi as *const _ as isize)),
        )
    }
    .0
}

fn rebuild(h: Handles, entries: &[Entry], roots: &HashSet<String>) {
    let (expanded, selected) = MODEL.with(|m| {
        let m = m.borrow();
        (m.expanded.clone(), m.selected.clone())
    });
    REBUILDING.with(|r| r.set(true));
    // SAFETY: our own tree control; REBUILDING makes the window procedure ignore its notifications.
    unsafe {
        SendMessageW(h.tree, WM_SETREDRAW, Some(WPARAM(0)), None);
        SendMessageW(h.tree, TVM_DELETEITEM, None, Some(LPARAM(TVI_ROOT.0)));
    }
    let mut items: Vec<Item> = Vec::with_capacity(entries.len());
    let mut by_path: HashMap<&str, isize> = HashMap::new();
    for e in entries {
        let parent = e
            .path
            .rsplit_once('/')
            .map_or(TVI_ROOT.0, |(p, _)| *by_path.get(p).unwrap_or(&TVI_ROOT.0));
        let handle = insert_item(h.tree, parent, &label(e, roots), items.len() as isize + 1);
        by_path.insert(&e.path, handle);
        items.push(Item {
            path: e.path.clone(),
            kind: e.kind,
            handle,
        });
    }
    // SAFETY: as above.
    unsafe {
        for it in items
            .iter()
            .filter(|i| i.kind == Kind::Dir && expanded.contains(&i.path))
        {
            SendMessageW(
                h.tree,
                TVM_EXPAND,
                Some(WPARAM(TVE_EXPAND.0 as usize)),
                Some(LPARAM(it.handle)),
            );
        }
        if let Some(sel) = selected
            .as_ref()
            .and_then(|s| items.iter().find(|i| &i.path == s))
        {
            SendMessageW(
                h.tree,
                TVM_SELECTITEM,
                Some(WPARAM(TVGN_CARET as usize)),
                Some(LPARAM(sel.handle)),
            );
        }
        SendMessageW(h.tree, WM_SETREDRAW, Some(WPARAM(1)), None);
        let _ = InvalidateRect(Some(h.tree), None, true);
    }
    REBUILDING.with(|r| r.set(false));
    MODEL.with(|m| {
        let mut m = m.borrow_mut();
        m.items = items;
        m.pending_new = None;
        if m.selected
            .as_ref()
            .is_some_and(|s| !m.items.iter().any(|i| &i.path == s))
        {
            m.selected = None;
        }
    });
}

fn item_for_param(param: isize) -> Option<(String, Kind, isize)> {
    MODEL.with(|m| {
        let m = m.borrow();
        let idx = usize::try_from(param - 1).ok()?;
        m.items.get(idx).map(|i| (i.path.clone(), i.kind, i.handle))
    })
}

/// Folder that "New folder", "Add files" etc. act on: the selected folder, or a file's folder.
fn target_dir() -> String {
    MODEL.with(|m| {
        let m = m.borrow();
        let Some(sel) = &m.selected else {
            return String::new();
        };
        match m.items.iter().find(|i| &i.path == sel).map(|i| i.kind) {
            Some(Kind::Dir) => sel.clone(),
            _ => sel
                .rsplit_once('/')
                .map_or(String::new(), |(p, _)| p.to_owned()),
        }
    })
}

fn begin_new_folder(h: Handles) {
    if MODEL.with(|m| m.borrow().pending_new.is_some()) || !MODEL.with(|m| m.borrow().can_edit) {
        return;
    }
    let parent = target_dir();
    let parent_handle = MODEL.with(|m| {
        m.borrow()
            .items
            .iter()
            .find(|i| i.path == parent)
            .map_or(TVI_ROOT.0, |i| i.handle)
    });
    // SAFETY: our own tree control.
    unsafe {
        if parent_handle != TVI_ROOT.0 {
            SendMessageW(
                h.tree,
                TVM_EXPAND,
                Some(WPARAM(TVE_EXPAND.0 as usize)),
                Some(LPARAM(parent_handle)),
            );
        }
        REBUILDING.with(|r| r.set(true));
        let handle = insert_item(h.tree, parent_handle, "New folder", 0);
        REBUILDING.with(|r| r.set(false));
        MODEL.with(|m| m.borrow_mut().pending_new = Some(PendingNew { parent, handle }));
        SendMessageW(h.tree, TVM_ENSUREVISIBLE, None, Some(LPARAM(handle)));
        SendMessageW(
            h.tree,
            TVM_SELECTITEM,
            Some(WPARAM(TVGN_CARET as usize)),
            Some(LPARAM(handle)),
        );
        SendMessageW(h.tree, TVM_EDITLABELW, None, Some(LPARAM(handle)));
    }
}

fn begin_rename(h: Handles) {
    let handle = MODEL.with(|m| {
        let m = m.borrow();
        let sel = m.selected.as_ref()?;
        m.items.iter().find(|i| &i.path == sel).map(|i| i.handle)
    });
    if let Some(handle) = handle {
        // SAFETY: our own tree control.
        unsafe { SendMessageW(h.tree, TVM_EDITLABELW, None, Some(LPARAM(handle))) };
    }
}

// ------------------------------------------------------------------ messages

fn enqueue(cmd: Command) {
    QUEUE.with(|q| q.borrow_mut().push_back(cmd));
    if let Some(h) = handles() {
        // SAFETY: our own window; posted so the work happens outside the current notification.
        let _ = unsafe { PostMessageW(Some(h.win), WM_QUEUE, WPARAM(0), LPARAM(0)) };
    }
}

fn button_clicked(h: Handles, id: usize) {
    let selected = MODEL.with(|m| m.borrow().selected.clone());
    match id {
        ID_NEW => begin_new_folder(h),
        ID_ADD_FILES => enqueue(Command::AddFiles { dest: target_dir() }),
        ID_ADD_FOLDER => enqueue(Command::AddFolder { dest: target_dir() }),
        ID_EXPORT => selected
            .map(|path| enqueue(Command::Export { path }))
            .unwrap_or_default(),
        ID_RENAME => begin_rename(h),
        ID_DELETE => selected
            .map(|path| enqueue(Command::Delete { path }))
            .unwrap_or_default(),
        _ => {}
    }
}

fn notify(h: Handles, nm: &NMHDR, lparam: LPARAM) -> LRESULT {
    if nm.hwndFrom != h.tree || REBUILDING.with(Cell::get) {
        return LRESULT(0);
    }
    // SAFETY: the notification structure matching `nm.code` is what the control sent.
    unsafe {
        match nm.code {
            TVN_SELCHANGEDW => {
                let t = &*(lparam.0 as *const NMTREEVIEWW);
                let selected = item_for_param(t.itemNew.lParam.0).map(|(p, _, _)| p);
                MODEL.with(|m| m.borrow_mut().selected = selected);
                update_chrome(h);
            }
            TVN_ITEMEXPANDEDW => {
                let t = &*(lparam.0 as *const NMTREEVIEWW);
                if let Some((path, _, _)) = item_for_param(t.itemNew.lParam.0) {
                    MODEL.with(|m| {
                        let mut m = m.borrow_mut();
                        if t.action == TVE_EXPAND {
                            m.expanded.insert(path);
                        } else if t.action == TVE_COLLAPSE {
                            m.expanded.remove(&path);
                        }
                    });
                }
            }
            TVN_BEGINLABELEDITW => {
                EDITING.with(|e| e.set(true));
                return LRESULT(0); // allow
            }
            TVN_ENDLABELEDITW => {
                EDITING.with(|e| e.set(false));
                let info = &*(lparam.0 as *const NMTVDISPINFOW);
                let new_text = from_wide_ptr(info.item.pszText)
                    .map(|s| s.trim().to_owned())
                    .filter(|s| !s.is_empty());
                let handle = info.item.hItem.0;
                let pending = MODEL.with(|m| m.borrow_mut().pending_new.take());
                match (pending, new_text) {
                    (Some(p), Some(name)) if p.handle == handle => enqueue(Command::NewFolder {
                        parent: p.parent,
                        name,
                    }),
                    (Some(_), _) => enqueue(Command::Refresh), // cancelled: drop the placeholder
                    (None, Some(name)) => {
                        if let Some((path, _, _)) =
                            item_for_param(info.item.lParam.0).or_else(|| {
                                // The control may not echo lParam; look the item up by handle instead.
                                MODEL.with(|m| {
                                    m.borrow()
                                        .items
                                        .iter()
                                        .find(|i| i.handle == handle)
                                        .map(|i| (i.path.clone(), i.kind, i.handle))
                                })
                            })
                            && path.rsplit('/').next() != Some(name.as_str())
                        {
                            enqueue(Command::Rename { path, name });
                        }
                    }
                    (None, None) => {}
                }
                return LRESULT(0); // we rebuild the tree ourselves
            }
            TVN_KEYDOWN => {
                let k = &*(lparam.0 as *const NMTVKEYDOWN);
                let selected = MODEL.with(|m| m.borrow().selected.clone());
                if k.wVKey == VK_DELETE.0 {
                    if let Some(path) = selected {
                        enqueue(Command::Delete { path });
                    }
                } else if k.wVKey == VK_F2.0 {
                    begin_rename(h);
                }
            }
            NM_DBLCLK => {}
            _ => {}
        }
    }
    LRESULT(0)
}

fn dropped(h: Handles, hdrop: HDROP) {
    let mut paths = Vec::new();
    // SAFETY: `hdrop` comes from WM_DROPFILES and is released with DragFinish.
    unsafe {
        let n = DragQueryFileW(hdrop, u32::MAX, None);
        for i in 0..n {
            let len = DragQueryFileW(hdrop, i, None) as usize;
            let mut buf = vec![0u16; len + 1];
            DragQueryFileW(hdrop, i, Some(&mut buf));
            paths.push(PathBuf::from(String::from_utf16_lossy(&buf[..len])));
        }
        DragFinish(hdrop);
    }
    let _ = h;
    if !paths.is_empty() {
        enqueue(Command::AddPaths {
            dest: target_dir(),
            paths,
        });
    }
}

unsafe extern "system" fn wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    let result = catch_unwind(AssertUnwindSafe(|| {
        handle_message(hwnd, msg, wparam, lparam)
    }));
    match result {
        Ok(Some(r)) => r,
        Ok(None) => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
        Err(_) => {
            crate::log::line(&format!(
                "panel: panic handling message {msg:#x} (swallowed)"
            ));
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
    }
}

fn handle_message(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> Option<LRESULT> {
    let h = handles().filter(|h| h.win == hwnd);
    match msg {
        WM_SIZE => {
            if let Some(h) = h {
                layout(&h);
            }
            Some(LRESULT(0))
        }
        WM_PAINT => {
            paint(hwnd);
            Some(LRESULT(0))
        }
        WM_GETMINMAXINFO => {
            // SAFETY: lparam points at a MINMAXINFO for this message.
            let mmi = unsafe { &mut *(lparam.0 as *mut MINMAXINFO) };
            mmi.ptMinTrackSize.x = dpi_scale(hwnd, 260);
            mmi.ptMinTrackSize.y = dpi_scale(hwnd, 300);
            Some(LRESULT(0))
        }
        WM_COMMAND => {
            let h = h?;
            let (id, code) = (wparam.0 & 0xFFFF, (wparam.0 >> 16) & 0xFFFF);
            if code == BN_CLICKED as usize {
                button_clicked(h, id);
            }
            Some(LRESULT(0))
        }
        WM_NOTIFY => {
            let h = h?;
            // SAFETY: WM_NOTIFY's lparam is an NMHDR*.
            let nm = unsafe { &*(lparam.0 as *const NMHDR) };
            Some(notify(h, nm, lparam))
        }
        WM_CTLCOLORSTATIC => {
            let h = h?;
            if HWND(lparam.0 as *mut c_void) != h.foot {
                return None;
            }
            let err = MODEL.with(|m| m.borrow().foot_is_error);
            let hdc = HDC(wparam.0 as *mut c_void);
            // SAFETY: the DC passed with this message.
            unsafe {
                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(
                    hdc,
                    if err {
                        COLORREF(0x001e26b3)
                    } else {
                        COLORREF(0x00202020)
                    },
                );
                Some(LRESULT(GetSysColorBrush(COLOR_3DFACE).0 as isize))
            }
        }
        WM_DROPFILES => {
            if let Some(h) = h {
                dropped(h, HDROP(wparam.0 as *mut c_void));
            }
            Some(LRESULT(0))
        }
        WM_QUEUE => {
            process_queue();
            Some(LRESULT(0))
        }
        WM_CLOSE => {
            // SAFETY: our own window; closing just hides it so it can be reopened.
            let _ = unsafe { ShowWindow(hwnd, SW_HIDE) };
            Some(LRESULT(0))
        }
        WM_DESTROY => {
            HANDLES.with(|c| c.set(None));
            None
        }
        _ => None,
    }
}

// ------------------------------------------------------------------ commands

fn adopt_authored(active: Option<&Active>) {
    let Some(path) = active.and_then(|a| a.path.as_deref()) else {
        return;
    };
    let digests = MODEL.with(|m| m.borrow().authored_unsaved.clone());
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
    let stale = MODEL.with(|m| {
        let m = m.borrow();
        force || m.key != active.key || m.live_read.is_none_or(|t| t.elapsed() > LIVE_REFRESH)
    });
    if stale {
        let read = live::read(&active.wb).ok().flatten();
        MODEL.with(|m| {
            let mut m = m.borrow_mut();
            m.live = read;
            m.live_read = Some(Instant::now());
            m.key = active.key.clone();
        });
    }
    MODEL.with(|m| m.borrow().live.clone())
}

fn process_queue() {
    let Some(h) = handles() else { return };
    let Some(app) = super::app() else {
        QUEUE.with(|q| q.borrow_mut().clear());
        return;
    };
    while let Some(cmd) = QUEUE.with(|q| q.borrow_mut().pop_front()) {
        let active = active_workbook(&app);
        let r = catch_unwind(AssertUnwindSafe(|| handle(h.win, active.as_ref(), cmd)));
        if r.is_err() {
            crate::log::line("panel: panic while handling a command (swallowed)");
            say(true, "Something went wrong; see addin.log.".to_owned());
        }
    }
    MODEL.with(|m| m.borrow_mut().signature.clear());
    tick(&app);
}

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

fn handle(owner: HWND, active: Option<&Active>, cmd: Command) {
    let Some(active) = active else {
        say(true, "Open a workbook first.".to_owned());
        return;
    };
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
        Command::AddPaths { dest, paths } => add(&current, dest, paths.clone()),
        Command::Export { path } => export(owner, &current, path),
    };
    match outcome {
        Outcome::Nothing => {}
        Outcome::Edit(bundle, msg) => commit(active, bundle, msg),
        Outcome::Note(n) => say(false, n),
        Outcome::Error(e) => say(true, e),
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
        crate::log::line(&format!("panel: writing the workbook failed: {e}"));
        say(true, format!("Could not update the workbook: {e}"));
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
            None => MODEL.with(|m| {
                let mut m = m.borrow_mut();
                m.authored_unsaved.push(digest);
                let n = m.authored_unsaved.len();
                if n > 8 {
                    m.authored_unsaved.drain(..n - 8);
                }
            }),
        }
    }
    MODEL.with(|m| {
        let mut m = m.borrow_mut();
        m.live = (!bundle.is_empty()).then_some(bundle);
        m.live_read = Some(Instant::now());
    });
    say(false, msg);
}

#[allow(dead_code)]
fn _status_is_used(_: Status) {}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::testutil::bundle_of;
    use windows::Win32::UI::Controls::TVM_GETCOUNT;
    use windows::Win32::UI::Input::KeyboardAndMouse::IsWindowEnabled;

    fn tree_count(h: Handles) -> isize {
        // SAFETY: our own tree control.
        unsafe { SendMessageW(h.tree, TVM_GETCOUNT, None, None) }.0
    }

    fn setup() -> Option<Handles> {
        shutdown();
        ensure_window(HWND::default())
    }

    fn show(h: Handles, b: &Bundle) {
        let entries = b.entries();
        let roots: HashSet<String> = b.root_scripts().iter().map(|f| f.path.clone()).collect();
        MODEL.with(|m| {
            let mut m = m.borrow_mut();
            m.can_edit = true;
            m.view = Some(ui::view(&Status::NoScripts));
            m.wb_name = "Book1.xlsx".into();
        });
        rebuild(h, &entries, &roots);
        update_chrome(h);
    }

    fn enabled(h: Handles, id: usize) -> bool {
        // SAFETY: our own child window.
        unsafe { IsWindowEnabled(GetDlgItem(Some(h.win), id as i32).unwrap()) }.as_bool()
    }

    #[test]
    fn the_window_renders_the_tree_and_turns_ui_events_into_commands() {
        let Some(h) = setup() else {
            eprintln!("cannot create windows on this machine: skipping panel test");
            return;
        };
        let b = bundle_of(&[
            ("main.lua", "x"),
            ("lib/util.lua", "y"),
            ("lib/deep/z.lua", "z"),
        ]);
        show(h, &b);
        assert_eq!(tree_count(h), 5); // lib, lib/deep, z.lua, util.lua, main.lua
        assert!(!enabled(h, ID_DELETE), "nothing selected yet");
        assert!(enabled(h, ID_NEW) && enabled(h, ID_ADD_FILES));

        // Selecting an item (a real TVN_SELCHANGED notification) enables the item buttons.
        let lib = MODEL
            .with(|m| {
                m.borrow()
                    .items
                    .iter()
                    .find(|i| i.path == "lib")
                    .map(|i| i.handle)
            })
            .unwrap();
        // SAFETY: our own tree control.
        unsafe {
            SendMessageW(
                h.tree,
                TVM_SELECTITEM,
                Some(WPARAM(TVGN_CARET as usize)),
                Some(LPARAM(lib)),
            )
        };
        assert_eq!(
            MODEL.with(|m| m.borrow().selected.clone()),
            Some("lib".into())
        );
        assert!(enabled(h, ID_DELETE) && enabled(h, ID_RENAME) && enabled(h, ID_EXPORT));
        assert_eq!(target_dir(), "lib");

        // Buttons queue commands for the selected folder.
        button_clicked(h, ID_ADD_FILES);
        button_clicked(h, ID_DELETE);
        let q: Vec<Command> = QUEUE.with(|q| q.borrow_mut().drain(..).collect());
        assert_eq!(
            q,
            [
                Command::AddFiles { dest: "lib".into() },
                Command::Delete { path: "lib".into() }
            ]
        );

        // Finishing an in-place rename arrives as TVN_ENDLABELEDIT and becomes a Rename command.
        let util = MODEL
            .with(|m| {
                m.borrow()
                    .items
                    .iter()
                    .find(|i| i.path == "lib/util.lua")
                    .map(|i| (i.handle, 4))
            })
            .unwrap();
        let mut name = wide("renamed.lua");
        let nm = NMTVDISPINFOW {
            hdr: NMHDR {
                hwndFrom: h.tree,
                idFrom: ID_TREE,
                code: TVN_ENDLABELEDITW,
            },
            item: TVITEMW {
                hItem: HTREEITEM(util.0),
                pszText: PWSTR(name.as_mut_ptr()),
                lParam: LPARAM(util.1),
                ..Default::default()
            },
        };
        // SAFETY: a stack NMTVDISPINFOW sent synchronously to our window.
        unsafe {
            SendMessageW(
                h.win,
                WM_NOTIFY,
                Some(WPARAM(ID_TREE)),
                Some(LPARAM(&nm as *const _ as isize)),
            )
        };
        let q: Vec<Command> = QUEUE.with(|q| q.borrow_mut().drain(..).collect());
        assert_eq!(
            q,
            [Command::Rename {
                path: "lib/util.lua".into(),
                name: "renamed.lua".into()
            }]
        );

        // A new folder starts as an editable placeholder; cancelling it removes it again.
        button_clicked(h, ID_NEW);
        assert_eq!(tree_count(h), 6);
        let pending = MODEL
            .with(|m| m.borrow().pending_new.as_ref().map(|p| p.handle))
            .unwrap();
        let cancel = NMTVDISPINFOW {
            hdr: NMHDR {
                hwndFrom: h.tree,
                idFrom: ID_TREE,
                code: TVN_ENDLABELEDITW,
            },
            item: TVITEMW {
                hItem: HTREEITEM(pending),
                ..Default::default()
            }, // pszText null = cancelled
        };
        // SAFETY: as above.
        unsafe {
            SendMessageW(
                h.win,
                WM_NOTIFY,
                Some(WPARAM(ID_TREE)),
                Some(LPARAM(&cancel as *const _ as isize)),
            )
        };
        let q: Vec<Command> = QUEUE.with(|q| q.borrow_mut().drain(..).collect());
        assert_eq!(q, [Command::Refresh]);
        // ...and a refresh rebuilds without the placeholder.
        show(h, &b);
        assert_eq!(tree_count(h), 5);
        shutdown();
    }

    #[test]
    fn labels_show_size_and_which_scripts_run() {
        let roots: HashSet<String> = ["main.lua".to_owned()].into();
        let e = |path: &str, kind, size| Entry {
            path: path.into(),
            kind,
            size,
        };
        assert_eq!(label(&e("lib", Kind::Dir, 0), &roots), "lib");
        assert!(label(&e("main.lua", Kind::File, 1536), &roots).contains("1.5 KB"));
        assert!(label(&e("main.lua", Kind::File, 10), &roots).contains("runs at load"));
        assert!(!label(&e("lib/x.lua", Kind::File, 10), &roots).contains("runs at load"));
    }
}
