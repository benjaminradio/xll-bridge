//! Ribbon button (COM add-in hosted by this same DLL).
//!
//! An XLL cannot contribute ribbon XML, so the DLL doubles as an in-proc COM server for an
//! Office COM add-in (`IDTExtensibility2` + `IRibbonExtensibility`). Because Excel loads that
//! add-in from the same file, it shares this module instance — and therefore the same
//! thread-local Lua runtime — with the worksheet functions.
//!
//! The button is one control whose icon/label follow the active workbook's trust state; a
//! 200 ms timer on Excel's UI thread notices workbook switches and invalidates the control.

#![allow(non_snake_case)] // COM method names

mod cells;
mod com;
mod dialogs;
mod dispatch;
mod icon;
mod live;
mod pane;

use std::cell::RefCell;
use std::ffi::c_void;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::PathBuf;

use windows::Win32::Foundation::{E_FAIL, E_POINTER, HWND, S_OK};
use windows::Win32::System::Com::{
    DISPATCH_FLAGS, DISPPARAMS, EXCEPINFO, IDispatch, IDispatch_Impl, IDispatch_Vtbl, ITypeInfo,
};
use windows::Win32::System::Variant::{VARIANT, VT_BYREF, VT_VARIANT};
use windows::Win32::UI::WindowsAndMessaging::{
    IDYES, KillTimer, MB_DEFBUTTON2, MB_ICONINFORMATION, MB_ICONWARNING, MB_OK, MB_SETFOREGROUND,
    MB_YESNO, MessageBoxW, SetTimer,
};
use windows::core::{BSTR, Error, GUID, HRESULT, Interface, PCWSTR, Result, implement, interface};

use super::with_runtime;
use crate::runtime::Status;
use crate::ui::{self, View};

pub use com::{register, unregister};

const BUTTON_ID: &str = "btnLuaStatus";
const POLL_MS: u32 = 200;

const RIBBON_XML: &str = r#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="tabXllBridge" label="Lua">
        <group id="grpXllBridge" label="Scripts">
          <button id="btnLuaStatus" size="large"
                  getLabel="GetLabel" getImage="GetImage"
                  getScreentip="GetScreentip" getSupertip="GetSupertip"
                  onAction="OnStatusClick"/>
          <button id="btnLuaRefresh" size="large" label="Re-run async" imageMso="RefreshAll"
                  screentip="Re-run async functions" supertip="Throw away the cached results of LUA.ASYNC in the active workbook and run those calls again."
                  onAction="OnRefreshClick"/>
          <button id="btnLuaFiles" size="large" label="Files" imageMso="FileOpen"
                  screentip="Lua files" supertip="Show or hide the panel that lists the files embedded in this workbook, with add, delete, rename, import and export."
                  onAction="OnFilesClick"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"#;

// IDispatch ids for the ribbon callbacks (names must match RIBBON_XML).
const CALLBACKS: [(&str, i32); 8] = [
    ("OnRibbonLoad", 1),
    ("GetLabel", 2),
    ("GetScreentip", 3),
    ("GetSupertip", 4),
    ("GetImage", 5),
    ("OnStatusClick", 6),
    ("OnFilesClick", 7),
    ("OnRefreshClick", 8),
];

#[interface("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
unsafe trait IDTExtensibility2: IDispatch {
    unsafe fn OnConnection(
        &self,
        application: *mut c_void,
        connect_mode: i32,
        add_in_inst: *mut c_void,
        custom: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn OnDisconnection(&self, remove_mode: i32, custom: *mut *mut c_void) -> HRESULT;
    unsafe fn OnAddInsUpdate(&self, custom: *mut *mut c_void) -> HRESULT;
    unsafe fn OnStartupComplete(&self, custom: *mut *mut c_void) -> HRESULT;
    unsafe fn OnBeginShutdown(&self, custom: *mut *mut c_void) -> HRESULT;
}

#[interface("000C0396-0000-0000-C000-000000000046")]
unsafe trait IRibbonExtensibility: IDispatch {
    unsafe fn GetCustomUI(&self, ribbon_id: *const u16, ribbon_xml: *mut *mut u16) -> HRESULT;
}

#[derive(Default)]
struct Shared {
    app: Option<IDispatch>,
    ribbon: Option<IDispatch>,
    timer: usize,
    shown: Option<View>,
}

/// Open workbooks by (lower-cased) name, with their real on-disk paths. Worksheet functions
/// cannot call back into Excel's object model while calculating, so the timer keeps this table
/// current for them.
#[derive(Default)]
struct Books {
    by_name: std::collections::HashMap<String, PathBuf>,
    last_count: i32,
    last_refresh: Option<std::time::Instant>,
    /// Cells of a newly seen workbook with embedded files may have been calculated before we
    /// could tell them where it lives; recalculate once when it appears.
    recalc_pending: bool,
}

// Cells waiting to be marked dirty and recalculated because a background call finished, with
// the number of failed attempts (Excel refuses COM calls while it is busy or editing a cell).
thread_local! {
    static DIRTY: RefCell<Vec<(crate::jobs::CellRef, u8)>> = const { RefCell::new(Vec::new()) };
}

const MAX_DIRTY_ATTEMPTS: u8 = 50;
const XL_CALCULATION_AUTOMATIC: i32 = -4105;

fn dirty_cell(app: &IDispatch, c: &crate::jobs::CellRef) -> bool {
    let attempt = || -> Result<()> {
        let books = IDispatch::try_from(&dispatch::get(app, "Workbooks")?)?;
        let wb = IDispatch::try_from(&dispatch::get_with(
            &books,
            "Item",
            &[VARIANT::from(c.book.as_str())],
        )?)?;
        let sheets = IDispatch::try_from(&dispatch::get(&wb, "Worksheets")?)?;
        let ws = IDispatch::try_from(&dispatch::get_with(
            &sheets,
            "Item",
            &[VARIANT::from(c.sheet.as_str())],
        )?)?;
        let range = IDispatch::try_from(&dispatch::get_with(
            &ws,
            "Range",
            &[VARIANT::from(c.a1().as_str())],
        )?)?;
        dispatch::call(&range, "Dirty", &[]).map(drop)
    };
    attempt().is_ok()
}

/// Recalculates the cells whose background calls have finished (or were asked to re-run).
fn pump_async(app: &IDispatch) {
    let fresh = crate::jobs::global().take_completed();
    DIRTY.with(|d| d.borrow_mut().extend(fresh.into_iter().map(|c| (c, 0))));
    let waiting = DIRTY.with(|d| std::mem::take(&mut *d.borrow_mut()));
    if waiting.is_empty() {
        return;
    }
    let mut retry = Vec::new();
    let mut marked = false;
    for (cell, attempts) in waiting {
        if dirty_cell(app, &cell) {
            marked = true;
        } else if attempts < MAX_DIRTY_ATTEMPTS {
            retry.push((cell, attempts + 1));
        }
    }
    DIRTY.with(|d| d.borrow_mut().extend(retry));
    // In manual calculation mode the cells are only marked; F9 picks them up.
    let automatic = dispatch::get(app, "Calculation")
        .ok()
        .and_then(|v| i32::try_from(&v).ok())
        == Some(XL_CALCULATION_AUTOMATIC);
    if marked && automatic {
        let _ = dispatch::call(app, "Calculate", &[]);
    }
}

/// The "Re-run async" button: forget this workbook's cached background results and run them again.
fn refresh_async(app: &IDispatch) {
    let (path, _) = current();
    let cells = path
        .as_deref()
        .map(|p| crate::jobs::global().refresh_workbook(p))
        .unwrap_or_default();
    if cells.is_empty() {
        message(
            excel_hwnd(app),
            "This workbook has no background (LUA.ASYNC) results to re-run.",
            "Lua",
            MB_OK | MB_ICONINFORMATION,
        );
        return;
    }
    DIRTY.with(|d| d.borrow_mut().extend(cells.into_iter().map(|c| (c, 0))));
    pump_async(app);
}

const BOOKS_REFRESH: std::time::Duration = std::time::Duration::from_secs(2);

thread_local! {
    static BOOKS: RefCell<Books> = RefCell::new(Books::default());
    static SHARED: RefCell<Shared> = RefCell::new(Shared::default());
}

/// The local path of the open workbook called `name` (e.g. `demo.xlsx`), if known.
pub(super) fn workbook_path_by_name(name: &str) -> Option<PathBuf> {
    BOOKS.with(|b| b.borrow().by_name.get(&name.to_lowercase()).cloned())
}

/// The name (e.g. `demo.xlsx`) under which Excel has the workbook at `path` open.
pub(super) fn workbook_name_by_path(path: &std::path::Path) -> Option<String> {
    let want = path.to_string_lossy().to_lowercase();
    BOOKS.with(|b| {
        b.borrow()
            .by_name
            .iter()
            .find(|(_, p)| p.to_string_lossy().to_lowercase() == want)
            .map(|(name, _)| name.clone())
    })
}

fn local_path_of(wb: &IDispatch) -> Option<PathBuf> {
    ["LocalFullName", "FullName"]
        .into_iter()
        .filter_map(|p| dispatch::get(wb, p).ok())
        .filter_map(|v| bstr_of(&v))
        .map(PathBuf::from)
        .find(|p| p.is_absolute() && p.is_file())
}

/// Re-reads the open-workbook table when the number of workbooks changed or it is a bit old.
fn refresh_books(app: &IDispatch) {
    let Some(coll) = dispatch::get(app, "Workbooks")
        .ok()
        .and_then(|v| IDispatch::try_from(&v).ok())
    else {
        return;
    };
    let count = dispatch::get(&coll, "Count")
        .ok()
        .and_then(|v| i32::try_from(&v).ok())
        .unwrap_or(0);
    let due = BOOKS.with(|b| {
        let b = b.borrow();
        b.last_count != count || b.last_refresh.is_none_or(|t| t.elapsed() > BOOKS_REFRESH)
    });
    if due {
        let mut map = std::collections::HashMap::new();
        for i in 1..=count {
            let Some(wb) = dispatch::get_with(&coll, "Item", &[VARIANT::from(i)])
                .ok()
                .and_then(|v| IDispatch::try_from(&v).ok())
            else {
                continue;
            };
            if let (Some(name), Some(path)) = (
                dispatch::get(&wb, "Name").ok().and_then(|v| bstr_of(&v)),
                local_path_of(&wb),
            ) {
                map.insert(name.to_lowercase(), path);
            }
        }
        let appeared: Vec<PathBuf> = BOOKS.with(|b| {
            let b = b.borrow();
            map.iter()
                .filter(|(k, _)| !b.by_name.contains_key(*k))
                .map(|(_, p)| p.clone())
                .collect()
        });
        // Workbooks that were closed: their background results are of no use any more.
        let vanished: Vec<PathBuf> = BOOKS.with(|b| {
            let b = b.borrow();
            b.by_name
                .iter()
                .filter(|(k, _)| !map.contains_key(*k))
                .map(|(_, p)| p.clone())
                .collect()
        });
        for gone in &vanished {
            crate::jobs::global().forget_workbook(gone);
        }
        let has_files = appeared
            .iter()
            .any(|p| with_runtime(|rt| !matches!(rt.status(p), Status::NoScripts)));
        BOOKS.with(|b| {
            let mut b = b.borrow_mut();
            b.by_name = map;
            b.last_count = count;
            b.last_refresh = Some(std::time::Instant::now());
            b.recalc_pending |= has_files;
        });
    }
    if BOOKS.with(|b| b.borrow().recalc_pending)
        && dispatch::call(app, "CalculateFull", &[]).is_ok()
    {
        BOOKS.with(|b| b.borrow_mut().recalc_pending = false);
    }
}

pub(super) fn is_connected() -> bool {
    SHARED.with(|s| s.borrow().app.is_some())
}

#[implement(IDTExtensibility2, IRibbonExtensibility)]
pub(super) struct Addin;

// ------------------------------------------------------------------ current state

fn app() -> Option<IDispatch> {
    SHARED.with(|s| s.borrow().app.clone())
}

fn bstr_of(v: &VARIANT) -> Option<String> {
    BSTR::try_from(v).ok().map(|b| b.to_string())
}

fn active_workbook_path(app: &IDispatch) -> Option<PathBuf> {
    let wb = IDispatch::try_from(&dispatch::get(app, "ActiveWorkbook").ok()?).ok()?;
    // LocalFullName resolves OneDrive/SharePoint-synced files to their on-disk copy.
    ["LocalFullName", "FullName"]
        .into_iter()
        .filter_map(|p| dispatch::get(&wb, p).ok())
        .filter_map(|v| bstr_of(&v))
        .map(PathBuf::from)
        .find(|p| p.is_absolute() && p.is_file())
}

/// (path, view) for the active workbook.
fn current() -> (Option<PathBuf>, View) {
    let path = app().and_then(|a| active_workbook_path(&a));
    let view = match &path {
        Some(p) => ui::view(&with_runtime(|rt| rt.status(p))),
        None => ui::view_without_file(),
    };
    (path, view)
}

fn invalidate_button() {
    let Some(ribbon) = SHARED.with(|s| s.borrow().ribbon.clone()) else {
        return;
    };
    let _ = dispatch::call(&ribbon, "InvalidateControl", &[VARIANT::from(BUTTON_ID)]);
}

unsafe extern "system" fn timer_proc(_: HWND, _: u32, _: usize, _: u32) {
    let _ = catch_unwind(AssertUnwindSafe(|| {
        let (_, view) = current();
        let changed = SHARED.with(|s| {
            let mut s = s.borrow_mut();
            let changed = s.shown.as_ref() != Some(&view);
            s.shown = Some(view);
            changed
        });
        if changed {
            invalidate_button();
        }
        if let Some(app) = app() {
            refresh_books(&app);
            pump_async(&app);
            crate::xlapi::pump_global();
            pane::tick(&app);
        }
    }));
}

// ------------------------------------------------------------------ button click

fn excel_hwnd(app: &IDispatch) -> HWND {
    dispatch::get(app, "Hwnd")
        .ok()
        .and_then(|v| i32::try_from(&v).ok())
        .map_or(HWND::default(), |h| HWND(h as isize as *mut c_void))
}

fn message(
    owner: HWND,
    text: &str,
    caption: &str,
    style: windows::Win32::UI::WindowsAndMessaging::MESSAGEBOX_STYLE,
) -> bool {
    let (t, c): (Vec<u16>, Vec<u16>) = (
        text.encode_utf16().chain(Some(0)).collect(),
        caption.encode_utf16().chain(Some(0)).collect(),
    );
    // SAFETY: both buffers are NUL-terminated and outlive the call.
    unsafe {
        MessageBoxW(
            Some(owner),
            PCWSTR(t.as_ptr()),
            PCWSTR(c.as_ptr()),
            style | MB_SETFOREGROUND,
        ) == IDYES
    }
}

fn on_click() {
    let Some(app) = app() else { return };
    let owner = excel_hwnd(&app);
    let (path, view) = current();
    let Some(path) = path else {
        message(owner, &view.supertip, "Lua", MB_OK | MB_ICONINFORMATION);
        return;
    };
    let status = with_runtime(|rt| rt.status(&path));
    match status {
        Status::Blocked { names } => {
            let text = format!(
                "\"{}\" contains {} Lua script(s):\n\n{}\n\nLua scripts can read and write files and run programs on this computer. \
                 Only enable them if you wrote or trust this workbook.\n\nEnable these scripts?",
                path.file_name()
                    .map(|n| n.to_string_lossy().into_owned())
                    .unwrap_or_default(),
                names.len(),
                names
                    .iter()
                    .map(|n| format!("  \u{2022} {n}"))
                    .collect::<Vec<_>>()
                    .join("\n"),
            );
            if !message(
                owner,
                &text,
                "Enable Lua scripts?",
                MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2,
            ) {
                return;
            }
            match with_runtime(|rt| rt.enable(&path)) {
                // Cells that were refused earlier are not dirty; force them to recompute.
                Ok(()) => drop(dispatch::call(&app, "CalculateFull", &[])),
                Err(e) => {
                    message(
                        owner,
                        &format!("Could not enable Lua: {e}"),
                        "Lua",
                        MB_OK | MB_ICONWARNING,
                    );
                }
            }
        }
        _ => {
            message(
                owner,
                &format!("{}\n\n{}", view.screentip, view.supertip),
                "Lua",
                MB_OK | MB_ICONINFORMATION,
            );
        }
    }
    invalidate_button();
}

// ------------------------------------------------------------------ COM plumbing

fn ret(v: VARIANT, pvarresult: *mut VARIANT, params: &DISPPARAMS) {
    // Excel accepts either a function-style return value or a trailing ByRef out-parameter.
    if !pvarresult.is_null() {
        // SAFETY: caller-provided result slot.
        unsafe { *pvarresult = v.clone() };
    }
    if params.cArgs >= 2 && !params.rgvarg.is_null() {
        // SAFETY: rgvarg has cArgs entries; entry 0 is the last declared parameter.
        unsafe {
            let first = &(*params.rgvarg).Anonymous.Anonymous;
            if first.vt.0 == VT_BYREF.0 | VT_VARIANT.0 {
                let target = first.Anonymous.pvarVal;
                if !target.is_null() {
                    *target = v;
                }
            }
        }
    }
}

impl IDispatch_Impl for Addin_Impl {
    fn GetTypeInfoCount(&self) -> Result<u32> {
        Ok(0)
    }

    fn GetTypeInfo(&self, _: u32, _: u32) -> Result<ITypeInfo> {
        Err(Error::from_hresult(HRESULT(0x8002_0003u32 as i32))) // DISP_E_UNKNOWNNAME-ish: no type info
    }

    fn GetIDsOfNames(
        &self,
        _: *const GUID,
        names: *const PCWSTR,
        count: u32,
        _: u32,
        ids: *mut i32,
    ) -> Result<()> {
        let mut missing = false;
        for i in 0..count as usize {
            // SAFETY: caller provides `count` names and `count` id slots.
            let (name, slot) =
                unsafe { ((*names.add(i)).to_string().unwrap_or_default(), ids.add(i)) };
            let id = (i == 0)
                .then(|| {
                    CALLBACKS
                        .iter()
                        .find(|(n, _)| n.eq_ignore_ascii_case(&name))
                        .map(|(_, id)| *id)
                })
                .flatten();
            // SAFETY: see above.
            unsafe { *slot = id.unwrap_or(-1) };
            missing |= id.is_none() && i == 0;
        }
        if missing {
            Err(Error::from_hresult(HRESULT(0x8002_0006u32 as i32)))
        } else {
            Ok(())
        } // DISP_E_UNKNOWNNAME
    }

    fn Invoke(
        &self,
        id: i32,
        _: *const GUID,
        _: u32,
        _: DISPATCH_FLAGS,
        params: *const DISPPARAMS,
        result: *mut VARIANT,
        _: *mut EXCEPINFO,
        _: *mut u32,
    ) -> Result<()> {
        catch_unwind(AssertUnwindSafe(|| -> Result<()> {
            if params.is_null() {
                return Err(Error::from_hresult(E_POINTER));
            }
            // SAFETY: non-null, supplied by the caller for this call.
            let params = unsafe { &*params };
            match id {
                1 => {
                    // OnRibbonLoad(IRibbonUI)
                    if params.cArgs >= 1 && !params.rgvarg.is_null() {
                        // SAFETY: at least one argument present.
                        let ribbon = IDispatch::try_from(unsafe { &*params.rgvarg }).ok();
                        SHARED.with(|s| s.borrow_mut().ribbon = ribbon);
                    }
                }
                2 => ret(VARIANT::from(current().1.label.as_str()), result, params),
                3 => ret(
                    VARIANT::from(current().1.screentip.as_str()),
                    result,
                    params,
                ),
                4 => ret(VARIANT::from(current().1.supertip.as_str()), result, params),
                5 => {
                    let pic = icon::picture(current().1.light)?;
                    ret(VARIANT::from(pic), result, params);
                }
                8 => {
                    crate::log::note("ribbon: Re-run async clicked");
                    if let Some(app) = app() {
                        refresh_async(&app);
                    }
                }
                6 => {
                    crate::log::note("ribbon: status button clicked");
                    on_click()
                }
                7 => {
                    crate::log::note("ribbon: Files button clicked");
                    if let Some(app) = app() {
                        pane::toggle(&app);
                    }
                }
                _ => return Err(Error::from_hresult(HRESULT(0x8002_0003u32 as i32))), // DISP_E_MEMBERNOTFOUND
            }
            Ok(())
        }))
        .unwrap_or_else(|_| Err(Error::from_hresult(E_FAIL)))
    }
}

impl IDTExtensibility2_Impl for Addin_Impl {
    unsafe fn OnConnection(
        &self,
        application: *mut c_void,
        _: i32,
        _: *mut c_void,
        _: *mut *mut c_void,
    ) -> HRESULT {
        // SAFETY: `application` is a borrowed IDispatch pointer from Excel; we AddRef via clone.
        let app = unsafe { IDispatch::from_raw_borrowed(&application) }.cloned();
        // SAFETY: plain thread timer on Excel's UI thread; `timer_proc` has the TIMERPROC ABI.
        let timer = unsafe { SetTimer(None, 0, POLL_MS, Some(timer_proc)) };
        crate::log::note("ribbon: connected to Excel");
        cells::install();
        SHARED.with(|s| {
            let mut s = s.borrow_mut();
            s.app = app;
            s.timer = timer;
        });
        S_OK
    }

    unsafe fn OnDisconnection(&self, _: i32, _: *mut *mut c_void) -> HRESULT {
        crate::log::note("ribbon: disconnecting");
        pane::shutdown();
        SHARED.with(|s| {
            let mut s = s.borrow_mut();
            if s.timer != 0 {
                // SAFETY: timer id returned by SetTimer on this thread.
                let _ = unsafe { KillTimer(None, s.timer) };
            }
            *s = Shared::default();
        });
        S_OK
    }

    unsafe fn OnAddInsUpdate(&self, _: *mut *mut c_void) -> HRESULT {
        S_OK
    }
    unsafe fn OnStartupComplete(&self, _: *mut *mut c_void) -> HRESULT {
        S_OK
    }
    unsafe fn OnBeginShutdown(&self, _: *mut *mut c_void) -> HRESULT {
        S_OK
    }
}

impl IRibbonExtensibility_Impl for Addin_Impl {
    unsafe fn GetCustomUI(&self, _: *const u16, ribbon_xml: *mut *mut u16) -> HRESULT {
        if ribbon_xml.is_null() {
            return E_POINTER;
        }
        // SAFETY: out-pointer provided by Excel; ownership of the BSTR passes to the caller.
        unsafe { *ribbon_xml = BSTR::from(RIBBON_XML).into_raw() as *mut u16 };
        S_OK
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx, IClassFactory};
    use windows::Win32::System::Registry::{
        HKEY, HKEY_CURRENT_USER, KEY_READ, RegCloseKey, RegOpenKeyExW,
    };
    use windows::core::IUnknown;

    fn create_addin() -> IUnknown {
        // SAFETY: standard COM bootstrap on the test thread.
        unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            let mut raw = std::ptr::null_mut();
            let hr = com::DllGetClassObject(&com::CLSID_ADDIN, &IClassFactory::IID, &mut raw);
            assert!(hr.is_ok(), "DllGetClassObject failed: {hr:?}");
            let factory = IClassFactory::from_raw(raw);
            factory
                .CreateInstance::<Option<&IUnknown>, IUnknown>(None)
                .expect("CreateInstance")
        }
    }

    #[test]
    fn object_implements_the_interfaces_excel_queries() {
        let obj = create_addin();
        assert!(obj.cast::<IDTExtensibility2>().is_ok());
        assert!(obj.cast::<IRibbonExtensibility>().is_ok());
        assert!(obj.cast::<IDispatch>().is_ok());
    }

    #[test]
    fn ribbon_xml_callbacks_all_resolve_through_idispatch() {
        let obj = create_addin();
        let ribbon: IRibbonExtensibility = obj.cast().unwrap();
        let mut raw: *mut u16 = std::ptr::null_mut();
        // SAFETY: out-param receives a BSTR we take ownership of below.
        unsafe { ribbon.GetCustomUI(std::ptr::null(), &mut raw) }
            .ok()
            .unwrap();
        // SAFETY: GetCustomUI returned an owned BSTR.
        let xml = unsafe { BSTR::from_raw(raw) }.to_string();
        assert!(xml.contains("<customUI") && xml.contains(BUTTON_ID));

        let disp: IDispatch = obj.cast().unwrap();
        let mut checked = 0;
        for attr in [
            "onLoad",
            "getLabel",
            "getImage",
            "getScreentip",
            "getSupertip",
            "onAction",
        ] {
            let needle = format!("{attr}=\"");
            for (at, _) in xml.match_indices(&needle) {
                let start = at + needle.len();
                let name = &xml[start..start + xml[start..].find('"').unwrap()];
                let id = CALLBACKS
                    .iter()
                    .find(|(n, _)| *n == name)
                    .unwrap_or_else(|| panic!("{name} not dispatchable"))
                    .1;
                let wide: Vec<u16> = name.encode_utf16().chain(Some(0)).collect();
                let mut got = 0i32;
                // SAFETY: one NUL-terminated name, one id slot.
                unsafe {
                    disp.GetIDsOfNames(&GUID::zeroed(), &PCWSTR(wide.as_ptr()), 1, 0x0400, &mut got)
                }
                .unwrap();
                assert_eq!(got, id, "{name}");
                checked += 1;
            }
        }
        assert_eq!(
            checked, 8,
            "onLoad + 5 on the status button + 1 on the files button"
        );
        assert!(xml.contains("onAction=\"OnFilesClick\""));
    }

    #[test]
    fn callbacks_return_label_and_picture() {
        let obj = create_addin();
        let disp: IDispatch = obj.cast().unwrap();
        let control = [VARIANT::default()]; // IRibbonControl is unused
        let label = dispatch::call(&disp, "GetLabel", &control).unwrap();
        assert_eq!(bstr_of(&label).as_deref(), Some("Lua: none")); // not connected => no workbook
        let image = dispatch::call(&disp, "GetImage", &control).unwrap();
        let pic = IDispatch::try_from(&image).expect("GetImage must return an IPictureDisp");
        let width = dispatch::get(&pic, "Width").expect("picture has a Width");
        assert!(i32::try_from(&width).unwrap_or(0) > 0);
        assert!(dispatch::call(&disp, "NoSuchCallback", &[]).is_err());
    }

    fn key_exists(subkey: &str) -> bool {
        let w: Vec<u16> = subkey.encode_utf16().chain(Some(0)).collect();
        let mut key = HKEY::default();
        // SAFETY: NUL-terminated name.
        let rc = unsafe {
            RegOpenKeyExW(
                HKEY_CURRENT_USER,
                PCWSTR(w.as_ptr()),
                None,
                KEY_READ,
                &mut key,
            )
        };
        if rc.is_ok() {
            // SAFETY: opened above.
            let _ = unsafe { RegCloseKey(key) };
        }
        rc.is_ok()
    }

    #[test]
    fn registration_round_trip() {
        let addin = format!(
            "Software\\Microsoft\\Office\\Excel\\Addins\\{}",
            com::PROGID
        );
        com::register().expect("register");
        assert!(key_exists(&addin));
        assert!(key_exists(&format!(
            "Software\\Classes\\CLSID\\{}\\InprocServer32",
            com::CLSID_STR
        )));
        com::unregister().expect("unregister");
        assert!(!key_exists(&addin));
    }
}
