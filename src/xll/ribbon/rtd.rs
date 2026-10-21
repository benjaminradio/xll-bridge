//! The real-time data (RTD) server behind `LUA.STREAM`.
//!
//! Excel's RTD mechanism: a worksheet function calls `RTD(progid, , topic)`; Excel creates the
//! COM server named by the ProgID (this DLL, registered under HKCU like the ribbon add-in),
//! calls `ConnectData` when a topic is first used and `DisconnectData` when no cell uses it any
//! more. When there is something new the server calls `UpdateNotify` on the callback Excel gave
//! it and Excel then collects the changed values with `RefreshData`.
//!
//! The streams themselves live in [`crate::streams`]. `UpdateNotify` is only ever called from
//! the UI-thread timer, so no COM pointer is shared between threads.

use std::cell::{Cell, RefCell};
use std::ffi::c_void;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Mutex;
use std::sync::atomic::{AtomicU32, Ordering};
use std::time::{Duration, Instant};

use windows::Win32::Foundation::{
    DISP_E_MEMBERNOTFOUND, E_FAIL, E_NOTIMPL, E_POINTER, S_OK, VARIANT_BOOL,
};
use windows::Win32::System::Com::{
    DISPATCH_FLAGS, DISPPARAMS, EXCEPINFO, IDispatch, IDispatch_Impl, IDispatch_Vtbl, ITypeInfo,
    SAFEARRAY, SAFEARRAYBOUND,
};
use windows::Win32::System::Ole::{SafeArrayAccessData, SafeArrayCreate, SafeArrayUnaccessData};
use windows::Win32::System::Variant::{
    VARENUM, VARIANT, VT_ARRAY, VT_BOOL, VT_BYREF, VT_DISPATCH, VT_I2, VT_I4, VT_UNKNOWN,
    VT_VARIANT,
};
use windows::core::{BSTR, Error, GUID, HRESULT, Interface, PCWSTR, Result, implement, interface};

use super::cells::variant_of;
use crate::value::{ErrorCode, Value};

pub const PROGID_RTD: &str = "XllBridge.Rtd";
pub(super) const CLSID_RTD: GUID = GUID::from_u128(0x9A4C2D18_5E7B_4F36_8C01_D2B6A3E5F974);
pub(super) const CLSID_RTD_STR: &str = "{9A4C2D18-5E7B-4F36-8C01-D2B6A3E5F974}";

/// How long to wait for Excel to collect an update before announcing it again.
const RENOTIFY: Duration = Duration::from_secs(2);

#[interface("EC0E6191-DB51-11D3-8F3E-00C04F3651B8")]
pub(super) unsafe trait IRtdServer: IDispatch {
    unsafe fn ServerStart(&self, callback: *mut c_void, result: *mut i32) -> HRESULT;
    unsafe fn ConnectData(
        &self,
        topic_id: i32,
        strings: *mut *mut SAFEARRAY,
        get_new_values: *mut VARIANT_BOOL,
        result: *mut VARIANT,
    ) -> HRESULT;
    unsafe fn RefreshData(&self, topic_count: *mut i32, result: *mut *mut SAFEARRAY) -> HRESULT;
    unsafe fn DisconnectData(&self, topic_id: i32) -> HRESULT;
    unsafe fn Heartbeat(&self, result: *mut i32) -> HRESULT;
    unsafe fn ServerTerminate(&self) -> HRESULT;
}

#[interface("A43788C1-D91B-11D3-8F39-00C04F3651B8")]
pub(super) unsafe trait IRTDUpdateEvent: IDispatch {
    unsafe fn UpdateNotify(&self) -> HRESULT;
    unsafe fn HeartbeatInterval(&self, result: *mut i32) -> HRESULT;
    unsafe fn SetHeartbeatInterval(&self, value: i32) -> HRESULT;
    unsafe fn Disconnect(&self) -> HRESULT;
}

static CREATED: AtomicU32 = AtomicU32::new(0);
static STARTED: AtomicU32 = AtomicU32::new(0);
static CONNECTS: AtomicU32 = AtomicU32::new(0);
static DISCONNECTS: AtomicU32 = AtomicU32::new(0);
static LAST_PROBLEM: Mutex<String> = Mutex::new(String::new());

fn problem(msg: &str) {
    crate::log::line(&format!("rtd: {msg}"));
    if let Ok(mut p) = LAST_PROBLEM.lock() {
        *p = msg.to_owned();
    }
}

/// Creates the server the way any COM client would (class lookup, DLL load, class factory,
/// `IRtdServer`) to tell "the registration is broken" from "Excel never asked". For `LUA.STATUS`.
pub(crate) fn self_test() -> String {
    use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance};
    // SAFETY: standard in-process creation on Excel's COM-initialised UI thread.
    let made: Result<IRtdServer> =
        unsafe { CoCreateInstance(&CLSID_RTD, None, CLSCTX_INPROC_SERVER) };
    match made {
        Ok(server) => {
            CREATED.fetch_sub(1, Ordering::Relaxed); // not Excel's: keep the counter honest
            drop(server);
            "self-test ok".to_owned()
        }
        Err(e) => format!("self-test FAILED: {e}"),
    }
}

/// What has happened so far, for `LUA.STATUS`: did Excel create the server, start it, and
/// connect any stream to it?
pub(crate) fn summary() -> String {
    let last = LAST_PROBLEM.lock().map(|p| p.clone()).unwrap_or_default();
    format!(
        "rtd: created {}, started {}, connected {}, disconnected {}, streams running {}{}",
        CREATED.load(Ordering::Relaxed),
        STARTED.load(Ordering::Relaxed),
        CONNECTS.load(Ordering::Relaxed),
        DISCONNECTS.load(Ordering::Relaxed),
        crate::streams::global().running(),
        if last.is_empty() {
            String::new()
        } else {
            format!(", last problem: {last}")
        }
    )
}

thread_local! {
    static CALLBACK: RefCell<Option<IRTDUpdateEvent>> = const { RefCell::new(None) };
    static NOTIFIED: Cell<Option<Instant>> = const { Cell::new(None) };
}

/// Called by the UI-thread timer: tells Excel when streams have new values.
pub(super) fn tick() {
    if !crate::streams::global().has_updates() {
        return;
    }
    if NOTIFIED.get().is_some_and(|t| t.elapsed() < RENOTIFY) {
        return;
    }
    CALLBACK.with(|cb| {
        if let Some(cb) = cb.borrow().as_ref() {
            // SAFETY: calling the callback Excel gave us, on the thread it was given on.
            if unsafe { cb.UpdateNotify() }.is_ok() {
                NOTIFIED.set(Some(Instant::now()));
            }
        }
    });
}

#[implement(IRtdServer)]
pub(super) struct Rtd;

impl Rtd {
    pub(super) fn new() -> Self {
        CREATED.fetch_add(1, Ordering::Relaxed);
        Rtd
    }
}

impl IDispatch_Impl for Rtd_Impl {
    fn GetTypeInfoCount(&self) -> Result<u32> {
        Ok(0)
    }
    fn GetTypeInfo(&self, _: u32, _: u32) -> Result<ITypeInfo> {
        Err(Error::from_hresult(E_NOTIMPL))
    }
    fn GetIDsOfNames(
        &self,
        _: *const GUID,
        names: *const PCWSTR,
        count: u32,
        _: u32,
        ids: *mut i32,
    ) -> Result<()> {
        // Excel may call the server late-bound (by name) instead of through the vtable.
        let mut missing = false;
        for i in 0..count as usize {
            // SAFETY: the caller provides `count` names and `count` id slots.
            let (name, slot) =
                unsafe { ((*names.add(i)).to_string().unwrap_or_default(), ids.add(i)) };
            let id = (i == 0).then(|| dispid_of(&name)).flatten();
            if id.is_none() && i == 0 {
                problem(&format!("Excel asked for an unknown member '{name}'"));
                missing = true;
            }
            // SAFETY: see above.
            unsafe { *slot = id.unwrap_or(-1) };
        }
        if missing {
            Err(Error::from_hresult(HRESULT(0x8002_0006u32 as i32))) // DISP_E_UNKNOWNNAME
        } else {
            Ok(())
        }
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
        if params.is_null() {
            return Err(Error::from_hresult(E_POINTER));
        }
        // SAFETY: non-null, supplied by the caller for this call.
        let params = unsafe { &*params };
        let hr = guarded(|| invoke(self, id, params, result));
        hr.ok()
    }
}

// ------------------------------------------------------------------ late-bound calls

// DISPIDs of IRtdServer in Excel's type library.
const ID_SERVER_START: i32 = 10;
const ID_CONNECT_DATA: i32 = 11;
const ID_REFRESH_DATA: i32 = 12;
const ID_DISCONNECT_DATA: i32 = 13;
const ID_HEARTBEAT: i32 = 14;
const ID_SERVER_TERMINATE: i32 = 15;

fn dispid_of(name: &str) -> Option<i32> {
    [
        ("ServerStart", ID_SERVER_START),
        ("ConnectData", ID_CONNECT_DATA),
        ("RefreshData", ID_REFRESH_DATA),
        ("DisconnectData", ID_DISCONNECT_DATA),
        ("Heartbeat", ID_HEARTBEAT),
        ("ServerTerminate", ID_SERVER_TERMINATE),
    ]
    .into_iter()
    .find(|(n, _)| n.eq_ignore_ascii_case(name))
    .map(|(_, id)| id)
}

/// The `i`th argument counted from the left (IDispatch passes them right to left).
fn arg(params: &DISPPARAMS, i: usize) -> Option<&VARIANT> {
    let n = params.cArgs as usize;
    if i >= n || params.rgvarg.is_null() {
        return None;
    }
    // SAFETY: rgvarg holds cArgs entries.
    Some(unsafe { &*params.rgvarg.add(n - 1 - i) })
}

/// A by-value or by-reference integer argument, and where it lives if it was by reference.
fn int_arg(v: &VARIANT) -> Option<(i32, Option<*mut i32>)> {
    let raw = v.vt().0;
    let by_ref = raw & VT_BYREF.0 != 0;
    // SAFETY: each branch reads the union member that the type tag says is active.
    unsafe {
        let a = &*v.Anonymous.Anonymous;
        match (raw & !VT_BYREF.0, by_ref) {
            (t, false) if t == VT_I4.0 => Some((a.Anonymous.lVal, None)),
            (t, true) if t == VT_I4.0 => {
                let p = a.Anonymous.plVal;
                (!p.is_null()).then(|| (*p, Some(p)))
            }
            (t, false) if t == VT_I2.0 => Some((i32::from(a.Anonymous.iVal), None)),
            _ => None,
        }
    }
}

fn set_int(result: *mut VARIANT, value: i32) {
    if !result.is_null() {
        // SAFETY: caller-provided result slot, holding no valid VARIANT yet.
        unsafe { result.write(VARIANT::from(value)) };
    }
}

fn invoke(this: &Rtd_Impl, id: i32, params: &DISPPARAMS, result: *mut VARIANT) -> HRESULT {
    let bad_args = HRESULT(0x8002_0005u32 as i32); // DISP_E_TYPEMISMATCH
    match id {
        ID_SERVER_START => {
            let Some(cb) = arg(params, 0) else {
                return bad_args;
            };
            // SAFETY: the VARIANT holds a COM pointer when its tag says so; we only borrow it.
            let raw = unsafe {
                let a = &*cb.Anonymous.Anonymous;
                match cb.vt() {
                    t if t == VT_DISPATCH => a.Anonymous.pdispVal.as_ref().map(|d| d.as_raw()),
                    t if t == VT_UNKNOWN => a.Anonymous.punkVal.as_ref().map(|d| d.as_raw()),
                    _ => None,
                }
            };
            let mut started = 0;
            // SAFETY: `raw` is a live interface pointer for the duration of the call.
            let hr = unsafe {
                IRtdServer_Impl::ServerStart(
                    this,
                    raw.unwrap_or(std::ptr::null_mut()),
                    &mut started,
                )
            };
            set_int(result, started);
            hr
        }
        ID_CONNECT_DATA => {
            let (Some(topic), Some(strings), gnv) =
                (arg(params, 0), arg(params, 1), arg(params, 2))
            else {
                return bad_args;
            };
            let Some((topic, _)) = int_arg(topic) else {
                return bad_args;
            };
            // The array arrives by reference (SAFEARRAY**) or by value (SAFEARRAY*).
            // SAFETY: reads the member the type tag says is active.
            let mut sa: *mut SAFEARRAY = unsafe {
                let a = &*strings.Anonymous.Anonymous;
                if strings.vt().0 & VT_BYREF.0 != 0 {
                    let pp = a.Anonymous.pparray;
                    if pp.is_null() {
                        std::ptr::null_mut()
                    } else {
                        *pp
                    }
                } else {
                    a.Anonymous.parray
                }
            };
            let mut flag = VARIANT_BOOL(0);
            // SAFETY: as above; a by-reference VT_BOOL points at the caller's flag.
            let flag_ptr: *mut VARIANT_BOOL = gnv
                .and_then(|g| unsafe {
                    let a = &*g.Anonymous.Anonymous;
                    (g.vt().0 == VT_BYREF.0 | VT_BOOL.0).then_some(a.Anonymous.pboolVal)
                })
                .filter(|p| !p.is_null())
                .unwrap_or(&mut flag);
            let mut local = VARIANT::default();
            let out = if result.is_null() {
                &mut local as *mut VARIANT
            } else {
                result
            };
            // SAFETY: all pointers are valid for the call.
            unsafe { IRtdServer_Impl::ConnectData(this, topic, &mut sa, flag_ptr, out) }
        }
        ID_REFRESH_DATA => {
            let count = arg(params, 0).and_then(int_arg).and_then(|(_, p)| p);
            let mut local = 0;
            let mut sa: *mut SAFEARRAY = std::ptr::null_mut();
            // SAFETY: valid out pointers.
            let hr =
                unsafe { IRtdServer_Impl::RefreshData(this, count.unwrap_or(&mut local), &mut sa) };
            if hr.is_ok() && !result.is_null() {
                let mut out = VARIANT::default();
                // SAFETY: filling a fresh VARIANT with the array we own; it takes ownership.
                unsafe {
                    let inner = &mut *out.Anonymous.Anonymous;
                    inner.vt = VARENUM(VT_ARRAY.0 | VT_VARIANT.0);
                    inner.Anonymous.parray = sa;
                    result.write(out);
                }
            }
            hr
        }
        ID_DISCONNECT_DATA => {
            let Some((topic, _)) = arg(params, 0).and_then(int_arg) else {
                return bad_args;
            };
            // SAFETY: plain call.
            unsafe { IRtdServer_Impl::DisconnectData(this, topic) }
        }
        ID_HEARTBEAT => {
            let mut beat = 0;
            // SAFETY: valid out pointer.
            let hr = unsafe { IRtdServer_Impl::Heartbeat(this, &mut beat) };
            set_int(result, beat);
            hr
        }
        ID_SERVER_TERMINATE => {
            // SAFETY: plain call.
            unsafe { IRtdServer_Impl::ServerTerminate(this) }
        }
        _ => DISP_E_MEMBERNOTFOUND,
    }
}

/// The first string of the topic array Excel passes to `ConnectData`.
///
/// # Safety
/// `sa` must be null or a valid one-dimensional SAFEARRAY of VARIANT.
unsafe fn first_topic(sa: *mut SAFEARRAY) -> Option<String> {
    if sa.is_null() {
        return None;
    }
    let mut data = std::ptr::null_mut();
    // SAFETY: caller guarantees a valid array; the data pointer is only used until Unaccess.
    unsafe {
        SafeArrayAccessData(sa, &mut data).ok()?;
        let text = BSTR::try_from(&*(data as *const VARIANT))
            .ok()
            .map(|b| b.to_string());
        let _ = SafeArrayUnaccessData(sa);
        text
    }
}

/// Wraps a body so a panic becomes `E_FAIL` instead of unwinding into Excel.
fn guarded(f: impl FnOnce() -> HRESULT) -> HRESULT {
    catch_unwind(AssertUnwindSafe(f)).unwrap_or(E_FAIL)
}

impl IRtdServer_Impl for Rtd_Impl {
    unsafe fn ServerStart(&self, callback: *mut c_void, result: *mut i32) -> HRESULT {
        guarded(|| {
            if result.is_null() {
                return E_POINTER;
            }
            // SAFETY: Excel passes a valid IRTDUpdateEvent pointer; we take our own reference.
            let cb = unsafe { IRTDUpdateEvent::from_raw_borrowed(&callback) }.cloned();
            let ok = cb.is_some();
            CALLBACK.with(|c| *c.borrow_mut() = cb);
            NOTIFIED.set(None);
            STARTED.fetch_add(1, Ordering::Relaxed);
            if !ok {
                problem("ServerStart got no callback object");
            }
            // SAFETY: caller-provided out pointer.
            unsafe { result.write(i32::from(ok)) };
            S_OK
        })
    }

    unsafe fn ConnectData(
        &self,
        topic_id: i32,
        strings: *mut *mut SAFEARRAY,
        get_new_values: *mut VARIANT_BOOL,
        result: *mut VARIANT,
    ) -> HRESULT {
        guarded(|| {
            if strings.is_null() || result.is_null() {
                return E_POINTER;
            }
            // SAFETY: Excel passes a pointer to its topic array.
            let key = unsafe { first_topic(*strings) };
            let outcome = match key {
                Some(key) => crate::streams::global().connect(topic_id, &key),
                None => Err("no stream name".to_owned()),
            };
            CONNECTS.fetch_add(1, Ordering::Relaxed);
            let initial = match outcome {
                Ok(()) => Value::Err(ErrorCode::GettingData),
                Err(m) => {
                    problem(&format!("connect failed: {m}"));
                    Value::Str(format!("#LUA: {m}"))
                }
            };
            // SAFETY: caller-provided out pointers; `write` because they hold no valid VARIANT.
            unsafe {
                result.write(variant_of(&initial));
                if !get_new_values.is_null() {
                    get_new_values.write(VARIANT_BOOL(-1));
                }
            }
            S_OK
        })
    }

    unsafe fn RefreshData(&self, topic_count: *mut i32, result: *mut *mut SAFEARRAY) -> HRESULT {
        guarded(|| {
            if topic_count.is_null() || result.is_null() {
                return E_POINTER;
            }
            NOTIFIED.set(None);
            let updates = crate::streams::global().take_updates();
            // Excel expects object[2, count]: row 0 topic ids, row 1 values (0-based bounds).
            let bounds = [
                SAFEARRAYBOUND {
                    cElements: 2,
                    lLbound: 0,
                },
                SAFEARRAYBOUND {
                    cElements: updates.len().max(1) as u32,
                    lLbound: 0,
                },
            ];
            // SAFETY: the array is created zeroed with 2 x n VARIANT slots and filled by moving
            // freshly built VARIANTs in (column-major: row index varies fastest).
            unsafe {
                let sa = SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr());
                if sa.is_null() {
                    return E_FAIL;
                }
                let mut data = std::ptr::null_mut();
                if SafeArrayAccessData(sa, &mut data).is_err() {
                    return E_FAIL;
                }
                let items = data as *mut VARIANT;
                for (i, (id, value)) in updates.iter().enumerate() {
                    items.add(i * 2).write(VARIANT::from(*id));
                    items.add(i * 2 + 1).write(variant_of(value));
                }
                let _ = SafeArrayUnaccessData(sa);
                topic_count.write(updates.len() as i32);
                result.write(sa);
            }
            S_OK
        })
    }

    unsafe fn DisconnectData(&self, topic_id: i32) -> HRESULT {
        guarded(|| {
            crate::streams::global().disconnect(topic_id);
            DISCONNECTS.fetch_add(1, Ordering::Relaxed);
            S_OK
        })
    }

    unsafe fn Heartbeat(&self, result: *mut i32) -> HRESULT {
        if !result.is_null() {
            // SAFETY: caller-provided out pointer.
            unsafe { result.write(1) };
        }
        S_OK
    }

    unsafe fn ServerTerminate(&self) -> HRESULT {
        guarded(|| {
            crate::streams::global().terminate();
            CALLBACK.with(|c| *c.borrow_mut() = None);
            crate::log::note("rtd: server terminated");
            S_OK
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx};
    use windows::core::IUnknown;

    #[test]
    fn the_server_answers_the_interface_excel_asks_for() {
        // SAFETY: standard COM bootstrap on the test thread.
        let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };
        let unknown: IUnknown = Rtd::new().into();
        let server: IRtdServer = unknown.cast().unwrap();
        let mut ignored = 0;
        // SAFETY: valid out pointer.
        assert!(unsafe { server.Heartbeat(&mut ignored) }.is_ok());
        assert_eq!(ignored, 1);
        // Nothing new: an empty (but well-formed) answer.
        let (mut count, mut sa) = (-1, std::ptr::null_mut());
        // SAFETY: valid out pointers.
        assert!(unsafe { server.RefreshData(&mut count, &mut sa) }.is_ok());
        assert_eq!(count, 0);
        assert!(!sa.is_null());
    }

    #[test]
    fn late_bound_calls_reach_the_same_methods() {
        // SAFETY: standard COM bootstrap on the test thread.
        let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };
        let unknown: IUnknown = Rtd::new().into();
        let disp: IDispatch = unknown.cast().unwrap();
        let beat = crate::xll::ribbon::dispatch::call(&disp, "Heartbeat", &[]).unwrap();
        assert_eq!(i32::try_from(&beat).unwrap(), 1);
        let refreshed = crate::xll::ribbon::dispatch::call(&disp, "RefreshData", &[]).unwrap();
        assert_eq!(refreshed.vt().0, VT_ARRAY.0 | VT_VARIANT.0);
        assert!(crate::xll::ribbon::dispatch::call(&disp, "NoSuchThing", &[]).is_err());
        crate::xll::ribbon::dispatch::call(&disp, "ServerTerminate", &[]).unwrap();
    }
}
