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
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, COLOR_3DFACE, COLOR_3DSHADOW, CreateFontIndirectW, CreateSolidBrush,
    DEFAULT_GUI_FONT, DT_END_ELLIPSIS, DT_LEFT, DT_NOPREFIX, DT_SINGLELINE, DT_VCENTER,
    DeleteObject, DrawTextW, Ellipse, EndPaint, FillRect, GetObjectW, GetStockObject, GetSysColor,
    GetSysColorBrush, HDC, HFONT, HGDIOBJ, InvalidateRect, LOGFONTW, MapWindowPoints, NULL_PEN,
    PAINTSTRUCT, ScreenToClient, SelectObject, SetBkMode, SetTextColor, TRANSPARENT,
};
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{
    HTREEITEM, ICC_TREEVIEW_CLASSES, INITCOMMONCONTROLSEX, InitCommonControlsEx, NM_DBLCLK, NMHDR,
    NMTREEVIEWW, NMTVDISPINFOW, NMTVKEYDOWN, TVE_COLLAPSE, TVE_EXPAND, TVGN_CARET, TVGN_DROPHILITE,
    TVHITTESTINFO, TVI_LAST, TVI_ROOT, TVIF_PARAM, TVIF_TEXT, TVINSERTSTRUCTW, TVINSERTSTRUCTW_0,
    TVITEMW, TVM_DELETEITEM, TVM_EDITLABELW, TVM_ENSUREVISIBLE, TVM_EXPAND, TVM_HITTEST,
    TVM_INSERTITEMW, TVM_SELECTITEM, TVN_BEGINDRAGW, TVN_BEGINLABELEDITW, TVN_ENDLABELEDITW,
    TVN_ITEMEXPANDEDW, TVN_KEYDOWN, TVN_SELCHANGEDW, TVS_EDITLABELS, TVS_HASBUTTONS, TVS_HASLINES,
    TVS_LINESATROOT, TVS_SHOWSELALWAYS, WC_TREEVIEW,
};
use windows::Win32::UI::HiDpi::GetDpiForWindow;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetKeyState, ReleaseCapture, SetCapture, VK_CONTROL, VK_DELETE, VK_F2, VK_SHIFT,
};
use windows::Win32::UI::Shell::{
    DragAcceptFiles, DragFinish, DragQueryFileW, HDROP, ShellExecuteW,
};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{PCWSTR, PWSTR, w};

use super::{dialogs, dispatch, live, message};
use crate::dock::{Dock, Rect as DRect};
use crate::localfs;
use crate::panel::{self, Command};
use crate::runtime::Status;
use crate::settings::{self, Settings};
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
const ID_MODE: usize = 107;
const ID_CLOSE: usize = 108;
const ID_EDIT: usize = 109;
const ID_COPY: usize = 110;
/// Width of the strip on a docked panel's left edge that can be dragged to resize it (96 DPI).
const GRIP: i32 = 6;
const DOCK_TICK_MS: u32 = 50;
const BUTTONS: [(usize, PCWSTR); 8] = [
    (ID_NEW, w!("New folder")),
    (ID_ADD_FILES, w!("Add files\u{2026}")),
    (ID_ADD_FOLDER, w!("Add folder\u{2026}")),
    (ID_EXPORT, w!("Export\u{2026}")),
    (ID_RENAME, w!("Rename")),
    (ID_DELETE, w!("Delete")),
    (ID_COPY, w!("Copy over")),
    (ID_EDIT, w!("Open in editor")),
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
    /// The machine folder's contents and when they were last read.
    machine: Vec<Entry>,
    machine_read: Option<Instant>,
    roots_seeded: bool,
    message: Option<(bool, String, Instant)>,
    signature: String,
    authored_unsaved: Vec<String>,
    busy: bool,
    settings: Settings,
    settings_loaded: bool,
}

/// A docked panel's link to the Excel frame it lives in.
struct DockCtx {
    frame: HWND,
    desk: HWND,
    dock: Dock,
    width_px: i32,
    timer: usize,
    dragging: bool,
}

thread_local! {
    static DOCK: RefCell<Option<DockCtx>> = const { RefCell::new(None) };
    /// Whether the current window is a docked child (as opposed to a floating tool window).
    static WIN_DOCKED: Cell<bool> = const { Cell::new(false) };
    static HANDLES: Cell<Option<Handles>> = const { Cell::new(None) };
    static MODEL: RefCell<Model> = RefCell::new(Model::default());
    static QUEUE: RefCell<VecDeque<Command>> = const { RefCell::new(VecDeque::new()) };
    /// True while we rebuild the tree ourselves: its notifications must be ignored.
    static REBUILDING: Cell<bool> = const { Cell::new(false) };
    static EDITING: Cell<bool> = const { Cell::new(false) };
    /// The node being dragged (set by TVN_BEGINDRAG while the mouse is captured).
    static DRAG: RefCell<Option<String>> = const { RefCell::new(None) };
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
    dock_stop();
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

/// Creates the panel window (hidden). `parent` is Excel's main frame when `docked` (the panel is
/// then a child of it) or Excel's main window as owner when floating.
fn create_window(parent: HWND, docked: bool) -> Option<Handles> {
    register_class();
    // SAFETY: Win32 window creation on the UI thread; all strings are static or outlive the calls.
    unsafe {
        let _ = InitCommonControlsEx(&INITCOMMONCONTROLSEX {
            dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_TREEVIEW_CLASSES,
        });
        let hinst = GetModuleHandleW(None).ok()?;
        let have_parent = !parent.0.is_null();
        let w_px = dpi_scale(parent, PANEL_WIDTH);
        let win = if docked {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                CLASS,
                TITLE,
                WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
                0,
                0,
                w_px,
                dpi_scale(parent, 400),
                Some(parent),
                None,
                Some(hinst.into()),
                None,
            )
        } else {
            // Start flush against the right edge of Excel, below the ribbon.
            let mut rc = RECT::default();
            let have_rect = have_parent && GetWindowRect(parent, &mut rc).is_ok();
            let (x, y, h) = if have_rect {
                let y = rc.top + dpi_scale(parent, 140);
                (
                    rc.right - w_px - dpi_scale(parent, 14),
                    y,
                    (rc.bottom - y - dpi_scale(parent, 60)).max(dpi_scale(parent, 360)),
                )
            } else {
                (100, 100, dpi_scale(parent, 520))
            };
            CreateWindowExW(
                WS_EX_TOOLWINDOW,
                CLASS,
                TITLE,
                WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_CLIPCHILDREN,
                x,
                y,
                w_px,
                h,
                have_parent.then_some(parent),
                None,
                Some(hinst.into()),
                None,
            )
        }
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
        let button = |id: usize, label: PCWSTR| {
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
        };
        for (id, label) in BUTTONS {
            button(id, label);
        }
        button(ID_MODE, if docked { w!("Float") } else { w!("Dock") });
        button(ID_CLOSE, w!("\u{2715}"));
        SendMessageW(
            foot,
            WM_SETFONT,
            Some(WPARAM(normal.0 as usize)),
            Some(LPARAM(1)),
        );
        DragAcceptFiles(win, true);
        let h = Handles { win, tree, foot };
        HANDLES.with(|c| c.set(Some(h)));
        WIN_DOCKED.with(|d| d.set(docked));
        layout(&h);
        Some(h)
    }
}

fn load_settings() -> Settings {
    MODEL.with(|m| {
        let mut m = m.borrow_mut();
        if !m.settings_loaded {
            m.settings = settings::load();
            m.settings_loaded = true;
        }
        m.settings
    })
}

/// Excel's frame window for the active window, and its workbook area (`XLDESK`).
/// In Excel 2013+ every workbook window has its own frame; older Excel has one.
fn find_frame(app: &IDispatch) -> Option<(HWND, HWND)> {
    let class_is = |h: HWND, name: &str| {
        let mut buf = [0u16; 64];
        // SAFETY: plain query into a local buffer.
        let n = unsafe { GetClassNameW(h, &mut buf) } as usize;
        String::from_utf16_lossy(&buf[..n]) == name
    };
    let window_hwnd = dispatch::get(app, "ActiveWindow")
        .ok()
        .and_then(|v| IDispatch::try_from(&v).ok())
        .and_then(|w| dispatch::get(&w, "Hwnd").ok())
        .and_then(|v| i32::try_from(&v).ok())
        .map(|h| HWND(h as isize as *mut c_void));
    let candidates = [window_hwnd, Some(super::excel_hwnd(app))];
    for frame in candidates.into_iter().flatten() {
        // SAFETY: window queries on handles that may be stale; failures are handled.
        unsafe {
            if frame.0.is_null() || !IsWindow(Some(frame)).as_bool() {
                continue;
            }
            let root = GetAncestor(frame, GA_ROOT);
            for f in [frame, root] {
                if f.0.is_null() || !class_is(f, "XLMAIN") {
                    continue;
                }
                if let Ok(desk) = FindWindowExW(Some(f), None, w!("XLDESK"), PCWSTR::null()) {
                    return Some((f, desk));
                }
            }
        }
    }
    None
}

/// The panel window for the current mode, creating (or recreating) it as needed.
fn ensure_window(app: &IDispatch) -> Option<Handles> {
    let want_docked = load_settings().docked;
    let frame = want_docked.then(|| find_frame(app)).flatten();
    let docked = frame.is_some();
    if let Some(h) = handles() {
        if WIN_DOCKED.with(Cell::get) == docked {
            return Some(h);
        }
        dock_stop();
        // SAFETY: our own window.
        let _ = unsafe { DestroyWindow(h.win) };
        HANDLES.with(|c| c.set(None));
    }
    if want_docked && !docked {
        crate::log::note("panel: could not find Excel's workbook area; floating instead");
    }
    create_window(
        frame.map_or_else(|| super::excel_hwnd(app), |(f, _)| f),
        docked,
    )
}

fn layout(h: &Handles) {
    // SAFETY: our own windows.
    unsafe {
        let mut rc = RECT::default();
        if GetClientRect(h.win, &mut rc).is_err() {
            return;
        }
        let s = |v: i32| dpi_scale(h.win, v);
        let lm = if WIN_DOCKED.with(Cell::get) {
            s(GRIP)
        } else {
            0
        };
        let (w, total_h) = (rc.right, rc.bottom);
        let (header, pad, btn_h, foot_h) = (s(46), s(6), s(26), s(24));
        // Header buttons, right-aligned.
        let (close_w, mode_w, hb_h) = (s(26), s(50), s(22));
        if let Ok(b) = GetDlgItem(Some(h.win), ID_CLOSE as i32) {
            let _ = MoveWindow(b, w - pad - close_w, s(8), close_w, hb_h, true);
        }
        if let Ok(b) = GetDlgItem(Some(h.win), ID_MODE as i32) {
            let _ = MoveWindow(
                b,
                w - pad - close_w - s(4) - mode_w,
                s(8),
                mode_w,
                hb_h,
                true,
            );
        }
        let bw = ((w - lm - 4 * pad) / 3).max(40);
        for (i, (id, _)) in BUTTONS.iter().enumerate() {
            let (col, row) = ((i % 3) as i32, (i / 3) as i32);
            if let Ok(b) = GetDlgItem(Some(h.win), *id as i32) {
                let _ = MoveWindow(
                    b,
                    lm + pad + col * (bw + pad),
                    header + pad + row * (btn_h + pad / 2),
                    bw,
                    btn_h,
                    true,
                );
            }
        }
        let rows = BUTTONS.len().div_ceil(3) as i32;
        let tree_top = header + pad * 2 + btn_h * rows + pad / 2 * (rows - 1);
        let _ = MoveWindow(
            h.tree,
            lm,
            tree_top,
            w - lm,
            (total_h - tree_top - foot_h).max(10),
            true,
        );
        let _ = MoveWindow(
            h.foot,
            lm + pad,
            total_h - foot_h + s(4),
            (w - lm - 2 * pad).max(10),
            foot_h - s(4),
            true,
        );
    }
}

// ------------------------------------------------------------------ show / hide

pub fn toggle(app: &IDispatch) {
    let owner = super::excel_hwnd(app);
    let Some(h) = ensure_window(app) else {
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
    if unsafe { IsWindowVisible(h.win) }.as_bool() {
        hide_panel(h);
    } else {
        show_panel(app, h);
        tick(app);
    }
}

fn hide_panel(h: Handles) {
    dock_stop();
    // SAFETY: our own window.
    let _ = unsafe { ShowWindow(h.win, SW_HIDE) };
}

fn show_panel(app: &IDispatch, h: Handles) {
    MODEL.with(|m| m.borrow_mut().signature.clear());
    if WIN_DOCKED.with(Cell::get) {
        match find_frame(app) {
            Some((frame, desk)) => dock_start(h, frame, desk),
            None => {
                crate::log::line(
                    "panel: lost Excel's workbook area while opening; showing floating",
                );
                // SAFETY: our own window.
                let _ = unsafe { ShowWindow(h.win, SW_SHOW) };
            }
        }
    } else {
        // SAFETY: our own window.
        let _ = unsafe { ShowWindow(h.win, SW_SHOW) };
    }
}

// ------------------------------------------------------------------ docking into Excel's frame

fn rect_in(frame: HWND, w: HWND) -> Option<DRect> {
    let mut r = RECT::default();
    // SAFETY: window geometry queries.
    unsafe {
        GetWindowRect(w, &mut r).ok()?;
        let mut pts = [
            POINT {
                x: r.left,
                y: r.top,
            },
            POINT {
                x: r.right,
                y: r.bottom,
            },
        ];
        MapWindowPoints(None, Some(frame), &mut pts);
        Some(DRect {
            left: pts[0].x,
            top: pts[0].y,
            right: pts[1].x,
            bottom: pts[1].y,
        })
    }
}

fn move_to(w: HWND, r: DRect) {
    // SAFETY: our own / Excel's windows; a failed move is harmless.
    let _ = unsafe {
        MoveWindow(
            w,
            r.left,
            r.top,
            (r.right - r.left).max(0),
            (r.bottom - r.top).max(0),
            true,
        )
    };
}

unsafe extern "system" fn dock_timer(_: HWND, _: u32, _: usize, _: u32) {
    let _ = catch_unwind(dock_step);
}

/// Parks the (child) panel in a strip carved off the right edge of Excel's workbook area and
/// starts the watchdog that keeps it there.
fn dock_start(h: Handles, frame: HWND, desk: HWND) {
    dock_stop();
    let width_px = dpi_scale(frame, load_settings().panel_width);
    // SAFETY: SetParent between two windows of this thread; the timer runs on this thread.
    let timer = unsafe {
        if GetParent(h.win).ok() != Some(frame) {
            let _ = SetParent(h.win, Some(frame));
        }
        let _ = ShowWindow(h.win, SW_SHOW);
        let _ = SetWindowPos(
            h.win,
            Some(HWND_TOP),
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        SetTimer(None, 0, DOCK_TICK_MS, Some(dock_timer))
    };
    DOCK.with(|d| {
        *d.borrow_mut() = Some(DockCtx {
            frame,
            desk,
            dock: Dock::new(),
            width_px,
            timer,
            dragging: false,
        })
    });
    dock_step();
}

/// One watchdog step: if Excel re-laid out its workbook area, carve the strip out again.
fn dock_step() {
    let Some(h) = handles() else { return };
    let Some((frame, desk, mut dock, width)) = DOCK.with(|d| {
        d.borrow()
            .as_ref()
            .map(|c| (c.frame, c.desk, c.dock, c.width_px))
    }) else {
        return;
    };
    // SAFETY: handle validity queries.
    let (alive, iconic) = unsafe {
        (
            IsWindow(Some(frame)).as_bool() && IsWindow(Some(desk)).as_bool(),
            IsIconic(frame).as_bool(),
        )
    };
    if !alive {
        dock_stop();
        return;
    }
    if iconic {
        return;
    }
    let mut client = RECT::default();
    // SAFETY: our reference frame.
    if unsafe { GetClientRect(frame, &mut client) }.is_err() {
        return;
    }
    let Some(now) = rect_in(frame, desk) else {
        return;
    };
    let plan = dock.plan(now, client.right, width);
    if dock.is_fighting() {
        crate::log::line(
            "panel: Excel keeps resetting its workbook area; switching to a floating panel",
        );
        MODEL.with(|m| m.borrow_mut().settings.docked = false); // this session only; not saved
        enqueue(Command::Recreate);
        dock_stop();
        return;
    }
    if plan.desk != now {
        move_to(desk, plan.desk);
    }
    if rect_in(frame, h.win) != Some(plan.panel) {
        move_to(h.win, plan.panel);
    }
    DOCK.with(|d| {
        if let Some(c) = d.borrow_mut().as_mut() {
            c.dock = dock;
        }
    });
}

/// Gives Excel its workbook area back and forgets the frame. Safe to call when not docked.
fn dock_stop() {
    let Some(ctx) = DOCK.with(|d| d.borrow_mut().take()) else {
        return;
    };
    // SAFETY: kill our timer; query/restore windows that may already be gone.
    unsafe {
        let _ = KillTimer(None, ctx.timer);
        if IsWindow(Some(ctx.frame)).as_bool()
            && IsWindow(Some(ctx.desk)).as_bool()
            && let Some(now) = rect_in(ctx.frame, ctx.desk)
            && let Some(right) = ctx.dock.restore_right(now.right)
        {
            move_to(ctx.desk, DRect { right, ..now });
        }
    }
}

/// While docked, follows the active workbook window if it lives in another Excel frame.
fn follow_active_frame(app: &IDispatch, h: Handles) {
    let Some((frame, desk)) = find_frame(app) else {
        return;
    };
    let current = DOCK.with(|d| d.borrow().as_ref().map(|c| c.frame));
    if current.is_some() && current != Some(frame) {
        dock_start(h, frame, desk);
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
    if WIN_DOCKED.with(Cell::get) {
        follow_active_frame(app, h);
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
                Some(p) => with_runtime(|rt| (ui::view(&rt.status(p)), rt.disk_content_digest(p))),
                None => (ui::view_without_file(), None),
            };
            let dirty = bundle.as_ref().map(Bundle::content_digest) != disk;
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
    let digest = bundle
        .as_ref()
        .map(Bundle::content_digest)
        .unwrap_or_default();
    let machine = refresh_machine(false);
    let machine_sig: String = machine
        .iter()
        .map(|e| format!("{}:{};", e.path, e.size))
        .collect();
    let signature = format!(
        "{key}|{digest}|{:?}|{}|{dirty}|{message:?}|{can_edit}|{machine_sig}",
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
    let nodes = build_nodes(bundle.as_ref(), &machine);
    MODEL.with(|m| {
        let mut m = m.borrow_mut();
        if m.key != key || !m.roots_seeded {
            m.expanded.clear();
            // Both trees start open.
            m.expanded.insert(String::new());
            m.expanded.insert(panel::MACHINE_ROOT.to_owned());
            m.roots_seeded = true;
            if m.key != key {
                m.selected = None;
            }
        }
        m.wb_name = name;
        m.view = Some(view);
        m.foot = foot;
        m.foot_is_error = foot_is_error;
        m.can_edit = can_edit;
        m.signature = signature;
    });
    rebuild(h, &nodes);
    update_chrome(h);
}

/// The machine folder's contents, re-read every few seconds (or when forced).
fn refresh_machine(force: bool) -> Vec<Entry> {
    let stale = MODEL.with(|m| {
        let m = m.borrow();
        force || m.machine_read.is_none_or(|t| t.elapsed() > LIVE_REFRESH)
    });
    if stale {
        let listing = crate::dbstore::local_dir()
            .and_then(|d| localfs::list(&d).ok())
            .unwrap_or_default();
        MODEL.with(|m| {
            let mut m = m.borrow_mut();
            m.machine = listing;
            m.machine_read = Some(Instant::now());
        });
    }
    MODEL.with(|m| m.borrow().machine.clone())
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
        let sel = selected.as_deref();
        let on_machine = sel.is_some_and(panel::is_machine);
        let on_node = sel.is_some_and(|s| !panel::is_root(s));
        let on_machine_file = on_node
            && on_machine
            && MODEL.with(|m| {
                m.borrow()
                    .items
                    .iter()
                    .any(|i| Some(i.path.as_str()) == sel && i.kind == Kind::File)
            });
        // The machine folder needs no open workbook; the workbook's tree does.
        let usable = can_edit || on_machine;
        for (id, _) in BUTTONS {
            let on = match id {
                ID_EXPORT | ID_RENAME | ID_DELETE | ID_COPY => usable && on_node,
                ID_EDIT => on_machine_file,
                _ => usable,
            };
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
        let lm = if WIN_DOCKED.with(Cell::get) {
            s(GRIP)
        } else {
            0
        };
        let _ = Ellipse(hdc, lm + s(10), s(16), lm + s(10) + d, s(16) + d);
        SelectObject(hdc, old_brush);
        SelectObject(hdc, old_pen);
        let _ = DeleteObject(HGDIOBJ(dot.0));

        SetBkMode(hdc, TRANSPARENT);
        let (normal, bold) = make_fonts();
        let line = |font: HFONT, color: COLORREF, top: i32, s_: &str| {
            let mut buf: Vec<u16> = s_.encode_utf16().collect();
            let mut r = RECT {
                left: lm + s(32),
                top,
                right: rc.right - s(96),
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
        if lm > 0 {
            // The drag strip on a docked panel's left edge.
            let line = RECT {
                left: s(2),
                top: 0,
                right: s(3),
                bottom: rc.bottom,
            };
            FillRect(hdc, &line, GetSysColorBrush(COLOR_3DSHADOW));
        }
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

/// One row of the tree: its node id (see [`panel::MACHINE_ROOT`]), kind and text.
struct Node {
    id: String,
    kind: Kind,
    label: String,
}

/// Both trees: ThisWorkbook (the embedded files) and Machine (the machine folder).
fn build_nodes(bundle: Option<&Bundle>, machine: &[Entry]) -> Vec<Node> {
    let mut nodes = Vec::new();
    let roots: HashSet<String> = bundle
        .map(|b| b.root_scripts().iter().map(|f| f.path.clone()).collect())
        .unwrap_or_default();
    let count = bundle.map_or(0, Bundle::file_count);
    nodes.push(Node {
        id: String::new(),
        kind: Kind::Dir,
        label: match bundle {
            Some(_) => format!(
                "ThisWorkbook   ({count} file{})",
                if count == 1 { "" } else { "s" }
            ),
            None => "ThisWorkbook   (no workbook)".to_owned(),
        },
    });
    for e in bundle.map(Bundle::entries).unwrap_or_default() {
        nodes.push(Node {
            label: label(&e, &roots),
            id: e.path,
            kind: e.kind,
        });
    }
    nodes.push(Node {
        id: panel::MACHINE_ROOT.to_owned(),
        kind: Kind::Dir,
        label: "Machine".to_owned(),
    });
    for e in machine {
        nodes.push(Node {
            label: label(e, &HashSet::new()),
            id: panel::machine_id(&e.path),
            kind: e.kind,
        });
    }
    nodes
}

fn rebuild(h: Handles, nodes: &[Node]) {
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
    let mut items: Vec<Item> = Vec::with_capacity(nodes.len());
    let mut by_id: HashMap<&str, isize> = HashMap::new();
    for n in nodes {
        let parent = panel::parent_id(&n.id)
            .and_then(|p| by_id.get(p.as_str()).copied())
            .unwrap_or(TVI_ROOT.0);
        let handle = insert_item(h.tree, parent, &n.label, items.len() as isize + 1);
        by_id.insert(&n.id, handle);
        items.push(Item {
            path: n.id.clone(),
            kind: n.kind,
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
            _ => panel::parent_id(sel).unwrap_or_default(),
        }
    })
}

fn begin_new_folder(h: Handles) {
    let machine_target = panel::is_machine(&target_dir());
    if MODEL.with(|m| m.borrow().pending_new.is_some())
        || !(machine_target || MODEL.with(|m| m.borrow().can_edit))
    {
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
        let sel = m.selected.as_ref().filter(|s| !panel::is_root(s))?;
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
            .filter(|s| !panel::is_root(s))
            .map(|path| enqueue(Command::Export { path }))
            .unwrap_or_default(),
        ID_RENAME => begin_rename(h),
        ID_COPY => {
            if let Some(from) = selected.filter(|s| !panel::is_root(s)) {
                // The other tree's root (use drag and drop to pick a folder).
                let to_dir = if panel::is_machine(&from) {
                    String::new()
                } else {
                    panel::MACHINE_ROOT.to_owned()
                };
                enqueue(Command::Transfer {
                    from,
                    to_dir,
                    move_: false,
                });
            }
        }
        ID_EDIT => {
            if let Some(path) = selected.filter(|s| panel::is_machine(s)) {
                enqueue(Command::OpenInEditor { path });
            }
        }
        ID_DELETE => selected
            .filter(|s| !panel::is_root(s))
            .map(|path| enqueue(Command::Delete { path }))
            .unwrap_or_default(),
        ID_MODE => enqueue(Command::ToggleDock),
        ID_CLOSE => enqueue(Command::Close),
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
                // The two roots keep their names.
                let info = &*(lparam.0 as *const NMTVDISPINFOW);
                if item_for_param(info.item.lParam.0).is_some_and(|(p, _, _)| panel::is_root(&p)) {
                    return LRESULT(1); // cancel
                }
                EDITING.with(|e| e.set(true));
                return LRESULT(0); // allow
            }
            TVN_BEGINDRAGW => {
                let t = &*(lparam.0 as *const NMTREEVIEWW);
                if let Some((id, _, _)) = item_for_param(t.itemNew.lParam.0)
                    && !panel::is_root(&id)
                {
                    DRAG.with(|d| *d.borrow_mut() = Some(id));
                    let _ = SetCapture(h.win);
                }
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
                            && panel::node_name(&path) != name.as_str()
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
                    if let Some(path) = selected.filter(|s| !panel::is_root(s)) {
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

// ------------------------------------------------------------------ drag and drop inside the tree

/// The tree node under a point given in the panel window's client coordinates.
fn node_at(h: Handles, lparam: LPARAM) -> Option<(String, Kind, isize)> {
    let mut pts = [POINT {
        x: (lparam.0 & 0xFFFF) as i16 as i32,
        y: ((lparam.0 >> 16) & 0xFFFF) as i16 as i32,
    }];
    // SAFETY: both windows are ours; the hit-test structure outlives the synchronous message.
    let hit = unsafe {
        MapWindowPoints(Some(h.win), Some(h.tree), &mut pts);
        let mut info = TVHITTESTINFO {
            pt: pts[0],
            ..Default::default()
        };
        SendMessageW(
            h.tree,
            TVM_HITTEST,
            None,
            Some(LPARAM(&mut info as *mut _ as isize)),
        );
        info.hItem.0
    };
    if hit == 0 {
        return None;
    }
    MODEL.with(|m| {
        m.borrow()
            .items
            .iter()
            .find(|i| i.handle == hit)
            .map(|i| (i.path.clone(), i.kind, i.handle))
    })
}

/// The folder a drop on `node` lands in: the folder itself, or a file's folder.
fn drop_folder(id: &str, kind: Kind) -> String {
    if kind == Kind::Dir {
        id.to_owned()
    } else {
        panel::parent_id(id).unwrap_or_default()
    }
}

fn set_drop_highlight(h: Handles, handle: isize) {
    // SAFETY: our own tree control.
    unsafe {
        SendMessageW(
            h.tree,
            TVM_SELECTITEM,
            Some(WPARAM(TVGN_DROPHILITE as usize)),
            Some(LPARAM(handle)),
        );
    }
}

fn drag_move(h: Handles, lparam: LPARAM) {
    let target = node_at(h, lparam).map(|(id, kind, handle)| (drop_folder(&id, kind), handle));
    let source = DRAG.with(|d| d.borrow().clone());
    // Highlight only where a drop would do something: not onto the node itself.
    match (source, target) {
        (Some(src), Some((folder, handle))) if folder != src => set_drop_highlight(h, handle),
        _ => set_drop_highlight(h, 0),
    }
}

fn drag_end(h: Handles) {
    DRAG.with(|d| *d.borrow_mut() = None);
    set_drop_highlight(h, 0);
    // SAFETY: releases the capture we took in TVN_BEGINDRAG (a no-op if it is already gone).
    unsafe {
        let _ = ReleaseCapture();
    }
}

fn drag_drop(h: Handles, lparam: LPARAM) {
    let source = DRAG.with(|d| d.borrow().clone());
    let target = node_at(h, lparam).map(|(id, kind, _)| drop_folder(&id, kind));
    drag_end(h);
    let (Some(from), Some(to_dir)) = (source, target) else {
        return;
    };
    if to_dir == from || panel::parent_id(&from).as_deref() == Some(to_dir.as_str()) {
        return; // dropped where it already is
    }
    // Within one tree a drop moves (Ctrl copies); between the trees it copies (Shift moves).
    let same_tree = panel::is_machine(&from) == panel::is_machine(&to_dir);
    // SAFETY: reads the keyboard state.
    let (ctrl, shift) = unsafe {
        (
            GetKeyState(i32::from(VK_CONTROL.0)) < 0,
            GetKeyState(i32::from(VK_SHIFT.0)) < 0,
        )
    };
    let move_ = if same_tree { !ctrl } else { shift };
    enqueue(Command::Transfer {
        from,
        to_dir,
        move_,
    });
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
        WM_SETCURSOR => {
            if !WIN_DOCKED.with(Cell::get) || (lparam.0 & 0xFFFF) as u32 != HTCLIENT {
                return None;
            }
            let dragging = DOCK.with(|d| d.borrow().as_ref().is_some_and(|c| c.dragging));
            let mut pt = POINT::default();
            // SAFETY: cursor queries for our own window.
            unsafe {
                let _ = GetCursorPos(&mut pt);
                let _ = ScreenToClient(hwnd, &mut pt);
                if dragging || pt.x < dpi_scale(hwnd, GRIP) {
                    SetCursor(LoadCursorW(None, IDC_SIZEWE).ok());
                    return Some(LRESULT(1));
                }
            }
            None
        }
        WM_LBUTTONDOWN => {
            let x = (lparam.0 & 0xFFFF) as i16 as i32;
            if WIN_DOCKED.with(Cell::get) && x < dpi_scale(hwnd, GRIP) {
                DOCK.with(|d| {
                    if let Some(c) = d.borrow_mut().as_mut() {
                        c.dragging = true;
                    }
                });
                // SAFETY: capture for our own window; released on button-up.
                unsafe { SetCapture(hwnd) };
            }
            Some(LRESULT(0))
        }
        WM_MOUSEMOVE if DRAG.with(|d| d.borrow().is_some()) => {
            if let Some(h) = h {
                drag_move(h, lparam);
            }
            Some(LRESULT(0))
        }
        WM_LBUTTONUP if DRAG.with(|d| d.borrow().is_some()) => {
            if let Some(h) = h {
                drag_drop(h, lparam);
            }
            Some(LRESULT(0))
        }
        WM_CAPTURECHANGED if DRAG.with(|d| d.borrow().is_some()) => {
            // Capture was taken away (another window, Escape in a dialog): cancel the drag.
            if let Some(h) = h {
                drag_end(h);
            }
            Some(LRESULT(0))
        }
        WM_MOUSEMOVE => {
            let dragging = DOCK.with(|d| d.borrow().as_ref().is_some_and(|c| c.dragging));
            if dragging {
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let mut rc = RECT::default();
                // SAFETY: our own window.
                let _ = unsafe { GetClientRect(hwnd, &mut rc) };
                // Dragging the left edge left by n pixels makes the panel n pixels wider.
                let width = (rc.right - x).max(0);
                DOCK.with(|d| {
                    if let Some(c) = d.borrow_mut().as_mut() {
                        c.width_px = width;
                    }
                });
                dock_step();
            }
            Some(LRESULT(0))
        }
        WM_LBUTTONUP | WM_CAPTURECHANGED => {
            let was_dragging = DOCK.with(|d| {
                d.borrow_mut()
                    .as_mut()
                    .is_some_and(|c| std::mem::replace(&mut c.dragging, false))
            });
            if was_dragging {
                let mut rc = RECT::default();
                // SAFETY: release capture; read the final width back from the window itself.
                unsafe {
                    let _ = ReleaseCapture();
                    let _ = GetClientRect(hwnd, &mut rc);
                }
                let dpi = dpi_scale(hwnd, 96).max(1);
                let saved = MODEL.with(|m| {
                    let mut m = m.borrow_mut();
                    m.settings.panel_width = rc.right * 96 / dpi;
                    m.settings
                });
                settings::save(&saved);
            }
            Some(LRESULT(0))
        }
        WM_DESTROY => {
            dock_stop();
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
    let Some(app) = super::app() else {
        QUEUE.with(|q| q.borrow_mut().clear());
        return;
    };
    while let Some(cmd) = QUEUE.with(|q| q.borrow_mut().pop_front()) {
        // Mode switches and closing may destroy the window, so look it up afresh every time.
        let Some(h) = handles() else {
            QUEUE.with(|q| q.borrow_mut().clear());
            return;
        };
        match cmd {
            Command::Close => {
                hide_panel(h);
                continue;
            }
            Command::ToggleDock => {
                switch_mode(&app);
                continue;
            }
            Command::Recreate => {
                recreate(&app);
                continue;
            }
            _ => {}
        }
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

/// Docked <-> floating: remember the choice, then recreate the window in the other mode.
fn switch_mode(app: &IDispatch) {
    let saved = MODEL.with(|m| {
        let mut m = m.borrow_mut();
        m.settings.docked = !m.settings.docked;
        m.settings
    });
    settings::save(&saved);
    recreate(app);
}

/// Destroys the window and builds it again in the configured mode, shown.
fn recreate(app: &IDispatch) {
    dock_stop();
    if let Some(h) = handles() {
        // SAFETY: our own window; nothing touches it after this returns to the message loop.
        let _ = unsafe { DestroyWindow(h.win) };
        HANDLES.with(|c| c.set(None));
    }
    if let Some(h) = ensure_window(app) {
        show_panel(app, h);
    }
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

/// Does the command touch the machine folder, or move things between the two trees?
fn involves_machine(cmd: &Command) -> bool {
    match cmd {
        Command::NewFolder { parent, .. } => panel::is_machine(parent),
        Command::AddFiles { dest }
        | Command::AddFolder { dest }
        | Command::AddPaths { dest, .. } => panel::is_machine(dest),
        Command::Export { path }
        | Command::Delete { path }
        | Command::Rename { path, .. }
        | Command::OpenInEditor { path } => panel::is_machine(path),
        Command::Transfer { .. } => true,
        _ => false,
    }
}

/// Opens a file with whatever the user has set up for it; Notepad when nothing is.
fn open_in_editor(owner: HWND, path: &Path) -> Result<String, String> {
    let file = wide(&path.to_string_lossy());
    // SAFETY: NUL-terminated strings that outlive the call.
    let r = unsafe {
        ShellExecuteW(
            Some(owner),
            w!("open"),
            PCWSTR(file.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        )
    };
    if r.0 as isize > 32 {
        return Ok(format!("Opened {}", path.display()));
    }
    std::process::Command::new("notepad.exe")
        .arg(path)
        .spawn()
        .map(|_| format!("Opened {} in Notepad", path.display()))
        .map_err(|e| format!("cannot open {}: {e}", path.display()))
}

fn export_machine(owner: HWND, root: &Path, rel: &str) -> Result<String, String> {
    let path = localfs::path_of(root, rel)?;
    let name = rel.rsplit('/').next().unwrap_or(rel);
    if path.is_dir() {
        match dialogs::pick_folder(owner, "Choose where to export the folder") {
            None => Ok(String::new()),
            Some(dest) => localfs::export_dir(root, rel, &dest),
        }
    } else {
        match dialogs::save_file(owner, name) {
            None => Ok(String::new()),
            Some(dest) => localfs::export_file(root, rel, &dest),
        }
    }
}

/// Copies or moves a node between the trees (or within the workbook's tree, which is the only
/// combination the machine folder itself does not cover).
fn transfer(
    active: Option<&Active>,
    root: &Path,
    from: &str,
    to_dir: &str,
    move_: bool,
) -> Result<String, String> {
    let need_workbook = || active.ok_or_else(|| "Open a workbook first.".to_owned());
    match (panel::is_machine(from), panel::is_machine(to_dir)) {
        (true, true) => localfs::transfer_within(
            root,
            panel::machine_rel(from),
            panel::machine_rel(to_dir),
            move_,
        ),
        (true, false) => {
            let active = need_workbook()?;
            let mut b = refresh_live(active, true).unwrap_or_default();
            let msg = localfs::copy_to_bundle(root, panel::machine_rel(from), &mut b, to_dir)?;
            commit(active, b, msg);
            if move_ {
                localfs::delete(root, panel::machine_rel(from))?;
            }
            Ok(String::new()) // commit has reported
        }
        (false, true) => {
            let active = need_workbook()?;
            let b = refresh_live(active, true).unwrap_or_default();
            let msg = localfs::copy_from_bundle(root, panel::machine_rel(to_dir), &b, from)?;
            if move_ {
                let mut without = b.clone();
                without.remove(from).map_err(|e| e.to_string())?;
                commit(
                    active,
                    without,
                    format!("Moved {from} to the machine folder"),
                );
                return Ok(String::new());
            }
            Ok(msg)
        }
        (false, false) => {
            let active = need_workbook()?;
            let mut b = refresh_live(active, true).unwrap_or_default();
            let target = panel::join(to_dir, panel::node_name(from));
            let msg = if move_ {
                b.rename(from, &target).map_err(|e| e.to_string())?;
                format!("Moved to {target}")
            } else {
                b.copy(from, &target).map_err(|e| e.to_string())?;
                format!("Copied to {target}")
            };
            commit(active, b, msg);
            Ok(String::new())
        }
    }
}

fn handle_machine(owner: HWND, active: Option<&Active>, cmd: Command) {
    let Some(root) = crate::dbstore::local_dir() else {
        say(true, "The machine folder is not available.".to_owned());
        return;
    };
    let add = |paths: Vec<PathBuf>, dest: &str| {
        if paths.is_empty() {
            Ok(String::new())
        } else {
            localfs::add_paths(&root, panel::machine_rel(dest), &paths)
        }
    };
    let result = match &cmd {
        Command::NewFolder { parent, name } => {
            localfs::new_folder(&root, panel::machine_rel(parent), name)
        }
        Command::Rename { path, name } => localfs::rename(&root, panel::machine_rel(path), name),
        Command::Delete { path } => {
            let rel = panel::machine_rel(path);
            let what = localfs::delete_summary(&root, rel);
            if message(
                owner,
                &format!("Delete {what} from the machine folder?\n\nThis cannot be undone."),
                "Delete",
                MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2,
            ) {
                localfs::delete(&root, rel)
            } else {
                Ok(String::new())
            }
        }
        Command::AddFiles { dest } => add(dialogs::pick_files(owner), dest),
        Command::AddFolder { dest } => add(dialogs::pick_folders(owner), dest),
        Command::AddPaths { dest, paths } => add(paths.clone(), dest),
        Command::Export { path } => export_machine(owner, &root, panel::machine_rel(path)),
        Command::OpenInEditor { path } => localfs::path_of(&root, panel::machine_rel(path))
            .and_then(|p| open_in_editor(owner, &p)),
        Command::Transfer {
            from,
            to_dir,
            move_,
        } => transfer(active, &root, from, to_dir, *move_),
        _ => Ok(String::new()),
    };
    refresh_machine(true);
    match result {
        Ok(m) if m.is_empty() => {}
        Ok(m) => say(false, m),
        Err(e) => say(true, e),
    }
}

fn handle(owner: HWND, active: Option<&Active>, cmd: Command) {
    // The machine folder and drag-and-drop between the trees need no special handling of the
    // workbook here (the transfer helpers ask for it when they do).
    if involves_machine(&cmd) {
        handle_machine(owner, active, cmd);
        return;
    }
    let Some(active) = active else {
        say(true, "Open a workbook first.".to_owned());
        return;
    };
    let current = refresh_live(active, true).unwrap_or_default();
    let outcome = match &cmd {
        Command::Refresh | Command::ToggleDock | Command::Close | Command::Recreate => {
            Outcome::Nothing
        }
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
        Command::Transfer { .. } | Command::OpenInEditor { .. } => Outcome::Nothing,
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
                if with_runtime(|rt| rt.note_edit(path, &digest)) {
                    // The user's own change: run it now, no need to save and reopen.
                    let app = dispatch::get(&active.wb, "Application")
                        .ok()
                        .and_then(|v| IDispatch::try_from(&v).ok());
                    if let Some(app) = app
                        && let Err(e) = super::folder::run_live(&app, path, &bundle)
                    {
                        msg.push_str(&format!(". Could not start the new code: {e}"));
                    }
                } else {
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
        create_window(HWND::default(), false)
    }

    fn show(h: Handles, b: &Bundle) {
        let nodes = build_nodes(Some(b), &[]);
        MODEL.with(|m| {
            let mut m = m.borrow_mut();
            m.can_edit = true;
            m.view = Some(ui::view(&Status::NoScripts));
            m.wb_name = "Book1.xlsx".into();
        });
        rebuild(h, &nodes);
        update_chrome(h);
    }

    fn enabled(h: Handles, id: usize) -> bool {
        // SAFETY: our own child window.
        unsafe { IsWindowEnabled(GetDlgItem(Some(h.win), id as i32).unwrap()) }.as_bool()
    }

    fn select(h: Handles, id: &str) {
        let handle = MODEL
            .with(|m| {
                m.borrow()
                    .items
                    .iter()
                    .find(|i| i.path == id)
                    .map(|i| i.handle)
            })
            .unwrap_or_else(|| panic!("no node {id:?}"));
        // SAFETY: our own tree control.
        unsafe {
            SendMessageW(
                h.tree,
                TVM_SELECTITEM,
                Some(WPARAM(TVGN_CARET as usize)),
                Some(LPARAM(handle)),
            )
        };
    }

    #[test]
    fn the_machine_tree_sits_beside_the_workbook_and_its_buttons_follow_the_selection() {
        let Some(h) = setup() else {
            eprintln!("cannot create windows on this machine: skipping panel test");
            return;
        };
        let b = bundle_of(&[("main.lua", "x")]);
        let machine = vec![
            Entry {
                path: "lib".into(),
                kind: Kind::Dir,
                size: 0,
            },
            Entry {
                path: "lib/x.lua".into(),
                kind: Kind::File,
                size: 3,
            },
        ];
        let nodes = build_nodes(Some(&b), &machine);
        MODEL.with(|m| {
            let mut m = m.borrow_mut();
            m.can_edit = true;
            m.view = Some(ui::view(&Status::NoScripts));
        });
        rebuild(h, &nodes);
        update_chrome(h);
        // ThisWorkbook + main.lua + Machine + lib + lib/x.lua
        assert_eq!(tree_count(h), 5);

        // The roots cannot be renamed, deleted, exported or copied; opening needs a machine file.
        select(h, "");
        assert!(enabled(h, ID_NEW) && enabled(h, ID_ADD_FILES));
        assert!(!enabled(h, ID_DELETE) && !enabled(h, ID_RENAME) && !enabled(h, ID_COPY));
        assert!(!enabled(h, ID_EDIT));
        select(h, panel::MACHINE_ROOT);
        assert!(!enabled(h, ID_DELETE) && !enabled(h, ID_EDIT));
        assert_eq!(target_dir(), panel::MACHINE_ROOT);

        // A workbook file can be copied over but not opened in an editor.
        select(h, "main.lua");
        assert!(enabled(h, ID_COPY) && !enabled(h, ID_EDIT));
        button_clicked(h, ID_COPY);
        // A machine file can be opened, and copies over to the workbook root.
        let file = panel::machine_id("lib/x.lua");
        select(h, &file);
        assert!(enabled(h, ID_EDIT) && enabled(h, ID_COPY) && enabled(h, ID_DELETE));
        assert_eq!(target_dir(), panel::machine_id("lib"));
        button_clicked(h, ID_COPY);
        button_clicked(h, ID_EDIT);
        let q: Vec<Command> = QUEUE.with(|q| q.borrow_mut().drain(..).collect());
        assert_eq!(
            q,
            [
                Command::Transfer {
                    from: "main.lua".into(),
                    to_dir: panel::MACHINE_ROOT.into(),
                    move_: false
                },
                Command::Transfer {
                    from: file.clone(),
                    to_dir: String::new(),
                    move_: false
                },
                Command::OpenInEditor { path: file }
            ]
        );
        assert!(involves_machine(&q[0]) && involves_machine(&q[2]));
        assert!(!involves_machine(&Command::Delete {
            path: "main.lua".into()
        }));

        // With no workbook open the machine tree still works.
        MODEL.with(|m| m.borrow_mut().can_edit = false);
        select(h, &panel::machine_id("lib"));
        update_chrome(h);
        assert!(enabled(h, ID_NEW) && enabled(h, ID_DELETE));
        select(h, "main.lua");
        update_chrome(h);
        assert!(!enabled(h, ID_NEW) && !enabled(h, ID_DELETE));
        shutdown();
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
        assert_eq!(tree_count(h), 7); // ThisWorkbook, lib, lib/deep, z.lua, util.lua, main.lua, Machine
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
                    .map(|i| (i.handle, 5))
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
        assert_eq!(tree_count(h), 8);
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
        assert_eq!(tree_count(h), 7);
        shutdown();
    }

    unsafe extern "system" fn test_proc(h: HWND, m: u32, w: WPARAM, l: LPARAM) -> LRESULT {
        // SAFETY: forwards to the default procedure unchanged.
        unsafe { DefWindowProcW(h, m, w, l) }
    }

    /// A stand-in for Excel's frame: a top-level window with an `XLDESK` child.
    fn fake_excel() -> Option<(HWND, HWND)> {
        // SAFETY: test-only window creation on the test thread.
        unsafe {
            let hinst = GetModuleHandleW(None).ok()?;
            for class in [w!("XllTestFrame"), w!("XLDESK")] {
                let wc = WNDCLASSW {
                    lpfnWndProc: Some(test_proc),
                    hInstance: hinst.into(),
                    lpszClassName: class,
                    ..Default::default()
                };
                RegisterClassW(&wc); // already-registered is fine
            }
            let frame = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("XllTestFrame"),
                w!("fake excel"),
                WS_OVERLAPPEDWINDOW,
                0,
                0,
                1000,
                700,
                None,
                None,
                Some(hinst.into()),
                None,
            )
            .ok()?;
            let mut rc = RECT::default();
            GetClientRect(frame, &mut rc).ok()?;
            let desk = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("XLDESK"),
                w!(""),
                WS_CHILD,
                0,
                100,
                rc.right,
                400,
                Some(frame),
                None,
                Some(hinst.into()),
                None,
            )
            .ok()?;
            Some((frame, desk))
        }
    }

    fn client_width(w: HWND) -> i32 {
        let mut rc = RECT::default();
        // SAFETY: test window.
        let _ = unsafe { GetClientRect(w, &mut rc) };
        rc.right
    }

    #[test]
    fn docking_carves_a_strip_and_keeps_it_when_excel_relays_out() {
        let Some((frame, desk)) = fake_excel() else {
            eprintln!("cannot create windows on this machine: skipping docking test");
            return;
        };
        shutdown();
        let h = create_window(frame, true).expect("docked panel");
        dock_start(h, frame, desk);

        let want = dpi_scale(frame, 340);
        let full = client_width(frame);
        let desk_rect = rect_in(frame, desk).unwrap();
        let panel = rect_in(frame, h.win).unwrap();
        assert_eq!(
            desk_rect.right,
            full - want,
            "workbook area gave up the strip"
        );
        assert_eq!(
            (panel.left, panel.right),
            (full - want, full),
            "panel sits in the strip"
        );
        assert_eq!((panel.top, panel.bottom), (desk_rect.top, desk_rect.bottom));

        // Excel re-lays out at full width and moves the top edge (formula bar grew): we follow.
        // SAFETY: test windows.
        unsafe { MoveWindow(desk, 0, 130, full, 380, true).unwrap() };
        dock_step();
        let (d, p) = (
            rect_in(frame, desk).unwrap(),
            rect_in(frame, h.win).unwrap(),
        );
        assert_eq!(
            (d.right, d.top, p.left, p.top),
            (full - want, 130, full - want, 130)
        );

        // The frame is resized and Excel lays out at the new width.
        // SAFETY: test windows.
        unsafe {
            SetWindowPos(
                frame,
                None,
                0,
                0,
                1300,
                700,
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE,
            )
            .unwrap()
        };
        let wide = client_width(frame);
        assert!(wide > full);
        // SAFETY: test windows.
        unsafe { MoveWindow(desk, 0, 130, wide, 380, true).unwrap() };
        dock_step();
        let (d, p) = (
            rect_in(frame, desk).unwrap(),
            rect_in(frame, h.win).unwrap(),
        );
        assert_eq!((d.right, p.left, p.right), (wide - want, wide - want, wide));

        // Dragging the grip changes the width; the strip stays glued to the right edge.
        DOCK.with(|c| c.borrow_mut().as_mut().unwrap().width_px = 420);
        dock_step();
        assert_eq!(rect_in(frame, desk).unwrap().right, wide - 420);
        assert_eq!(rect_in(frame, h.win).unwrap().right, wide);

        // Closing gives Excel its width back.
        dock_stop();
        assert_eq!(rect_in(frame, desk).unwrap().right, wide);
        assert!(DOCK.with(|c| c.borrow().is_none()));
        shutdown();
        // SAFETY: test window.
        let _ = unsafe { DestroyWindow(frame) };
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
