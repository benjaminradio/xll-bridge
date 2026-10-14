//! "Security warning" bar shown when a workbook carries Lua scripts that aren't trusted yet.
//!
//! An XLL has no access to Office's own message bar, so this is a borderless yellow popup
//! owned by (and tracking) the Excel main window, just below the ribbon. It runs on Excel's
//! UI thread — UDFs are registered thread-unsafe, so that's where they are called — which
//! means no extra message pump is needed.

use std::cell::RefCell;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};

use windows_sys::Win32::Foundation::{HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows_sys::Win32::Graphics::Gdi::{
    ClientToScreen, CreateSolidBrush, DEFAULT_GUI_FONT, GetStockObject, HBRUSH, HDC, SetBkColor,
    SetTextColor,
};
use windows_sys::Win32::System::LibraryLoader::GetModuleHandleW;
use windows_sys::Win32::UI::HiDpi::GetDpiForWindow;
use windows_sys::Win32::UI::WindowsAndMessaging::*;

use super::ffi::{Xloper12, excel12, xlGetHwnd, xlretSuccess};
use super::with_runtime;

/// `SS_CENTERIMAGE`: vertically centre the label text (not exported by windows-sys).
const SS_CENTERIMAGE: u32 = 0x200;
const ID_ENABLE: isize = 1;
const ID_DISMISS: isize = 2;
const TIMER_TRACK: usize = 1;
const BAR_HEIGHT: i32 = 34;
/// Distance from the top of Excel's client area to the bottom of ribbon + formula bar at 96 DPI.
const BELOW_RIBBON: i32 = 150;

const BG: u32 = 0x00_CC_F2_FF; // COLORREF is 0x00BBGGRR: RGB(255, 242, 204)

struct State {
    bar: HWND,
    label: HWND,
    owner: HWND,
    workbook: PathBuf,
}

thread_local! {
    static STATE: RefCell<Option<State>> = const { RefCell::new(None) };
}

static CLASS_REGISTERED: AtomicBool = AtomicBool::new(false);

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

fn excel_hwnd() -> Option<HWND> {
    let mut out = Xloper12::nil();
    if excel12(xlGetHwnd, Some(&mut out), &mut []) != xlretSuccess {
        return None;
    }
    // Excel returns the HWND as a 32-bit integer.
    Some(unsafe { out.val.w } as u32 as usize as HWND)
}

/// Shows (or keeps) the bar for `workbook`. No-op if it is already showing for it.
pub fn show(workbook: &Path, scripts: usize) {
    if STATE.with(|s| {
        s.borrow()
            .as_ref()
            .is_some_and(|st| st.workbook == workbook)
    }) {
        return;
    }
    dismiss();
    let Some(owner) = excel_hwnd() else { return };
    // SAFETY: Win32 window creation on the UI thread; all pointers outlive the calls.
    unsafe {
        let hinst = GetModuleHandleW(std::ptr::null());
        let class = wide("XllBridgeBanner");
        if !CLASS_REGISTERED.swap(true, Ordering::SeqCst) {
            let wc = WNDCLASSW {
                style: 0,
                lpfnWndProc: Some(wndproc),
                cbClsExtra: 0,
                cbWndExtra: 0,
                hInstance: hinst,
                hIcon: std::ptr::null_mut(),
                hCursor: LoadCursorW(std::ptr::null_mut(), IDC_ARROW),
                hbrBackground: bg_brush(),
                lpszMenuName: std::ptr::null(),
                lpszClassName: class.as_ptr(),
            };
            RegisterClassW(&wc);
        }
        let bar = CreateWindowExW(
            WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            class.as_ptr(),
            std::ptr::null(),
            WS_POPUP | WS_VISIBLE,
            0,
            0,
            600,
            BAR_HEIGHT,
            owner,
            std::ptr::null_mut(),
            hinst,
            std::ptr::null(),
        );
        if bar.is_null() {
            return;
        }
        let font = GetStockObject(DEFAULT_GUI_FONT);
        let text = format!(
            "SECURITY WARNING   {scripts} Lua script(s) in this workbook have been disabled."
        );
        let label = child(bar, "STATIC", &text, SS_CENTERIMAGE, 0, hinst);
        let enable = child(
            bar,
            "BUTTON",
            "Enable Lua",
            BS_PUSHBUTTON as u32,
            ID_ENABLE,
            hinst,
        );
        let dismiss = child(
            bar,
            "BUTTON",
            "\u{2715}",
            BS_PUSHBUTTON as u32,
            ID_DISMISS,
            hinst,
        );
        for h in [label, enable, dismiss] {
            SendMessageW(h, WM_SETFONT, font as WPARAM, 1);
        }
        STATE.with(|s| {
            *s.borrow_mut() = Some(State {
                bar,
                label,
                owner,
                workbook: workbook.to_path_buf(),
            })
        });
        SetTimer(bar, TIMER_TRACK, 250, None);
        reposition();
    }
}

pub fn dismiss() {
    if let Some(st) = STATE.with(|s| s.borrow_mut().take()) {
        // SAFETY: handle was created by us on this thread.
        unsafe {
            KillTimer(st.bar, TIMER_TRACK);
            DestroyWindow(st.bar);
        }
    }
}

unsafe fn child(
    parent: HWND,
    class: &str,
    text: &str,
    style: u32,
    id: isize,
    hinst: *mut core::ffi::c_void,
) -> HWND {
    let (c, t) = (wide(class), wide(text));
    unsafe {
        CreateWindowExW(
            0,
            c.as_ptr(),
            t.as_ptr(),
            WS_CHILD | WS_VISIBLE | style,
            0,
            0,
            10,
            10,
            parent,
            id as _,
            hinst,
            std::ptr::null(),
        )
    }
}

fn bg_brush() -> HBRUSH {
    thread_local! { static BRUSH: HBRUSH = unsafe { CreateSolidBrush(BG) }; }
    BRUSH.with(|b| *b)
}

/// Lays the bar out across the top of Excel's client area and positions its children.
fn reposition() {
    STATE.with(|s| {
        let s = s.borrow();
        let Some(st) = s.as_ref() else { return };
        // SAFETY: handles validated by IsWindow below.
        unsafe {
            if IsWindow(st.owner) == 0 {
                return;
            }
            let visible = IsIconic(st.owner) == 0;
            ShowWindow(st.bar, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE });
            if !visible {
                return;
            }
            let dpi = GetDpiForWindow(st.owner).max(96) as i32;
            let scale = |v: i32| v * dpi / 96;
            let mut client = RECT {
                left: 0,
                top: 0,
                right: 0,
                bottom: 0,
            };
            GetClientRect(st.owner, &mut client);
            let mut origin = POINT { x: 0, y: 0 };
            ClientToScreen(st.owner, &mut origin);
            let (w, h) = (client.right - client.left, scale(BAR_HEIGHT));
            SetWindowPos(
                st.bar,
                std::ptr::null_mut(),
                origin.x,
                origin.y + scale(BELOW_RIBBON),
                w,
                h,
                SWP_NOACTIVATE | SWP_NOZORDER,
            );

            let (btn_w, pad) = (scale(100), scale(8));
            let enable = GetDlgItem(st.bar, ID_ENABLE as i32);
            let dismiss = GetDlgItem(st.bar, ID_DISMISS as i32);
            let dismiss_w = scale(30);
            MoveWindow(
                dismiss,
                w - dismiss_w - pad,
                scale(4),
                dismiss_w,
                h - scale(8),
                1,
            );
            MoveWindow(
                enable,
                w - dismiss_w - btn_w - 2 * pad,
                scale(4),
                btn_w,
                h - scale(8),
                1,
            );
            MoveWindow(
                st.label,
                pad,
                0,
                (w - dismiss_w - btn_w - 4 * pad).max(0),
                h,
                1,
            );
        }
    });
}

unsafe extern "system" fn wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_TIMER if wparam == TIMER_TRACK => {
                reposition();
                0
            }
            WM_CTLCOLORSTATIC => {
                let hdc = wparam as HDC;
                SetBkColor(hdc, BG);
                SetTextColor(hdc, 0);
                bg_brush() as LRESULT
            }
            WM_COMMAND => {
                match (wparam & 0xFFFF) as isize {
                    ID_ENABLE => on_enable(),
                    ID_DISMISS => {
                        // Destroying the window from inside its own message handler is fine,
                        // but defer so the button finishes processing the click first.
                        PostMessageW(hwnd, WM_CLOSE, 0, 0);
                    }
                    _ => {}
                }
                0
            }
            WM_CLOSE => {
                dismiss();
                0
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}

fn on_enable() {
    let Some((wb, label, enable)) = STATE.with(|s| {
        s.borrow().as_ref().map(|st| {
            (st.workbook.clone(), st.label, unsafe {
                GetDlgItem(st.bar, ID_ENABLE as i32)
            })
        })
    }) else {
        return;
    };
    let result = with_runtime(|rt| rt.enable(&wb));
    let message = match result {
        Ok(()) => "Lua enabled for this workbook. Press Ctrl+Alt+F9 to recalculate.".to_owned(),
        Err(e) => format!("Could not enable Lua: {e}"),
    };
    let ok = message.starts_with("Lua enabled");
    let text = wide(&message);
    // SAFETY: valid handles owned by this thread.
    unsafe {
        SetWindowTextW(label, text.as_ptr());
        if ok {
            ShowWindow(enable, SW_HIDE);
        }
    }
}
