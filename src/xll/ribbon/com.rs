//! In-proc COM server plumbing: class factory, DLL exports, per-user registration.

use std::ffi::c_void;
use std::panic::catch_unwind;

use windows::Win32::Foundation::{
    CLASS_E_NOAGGREGATION, E_FAIL, E_POINTER, HMODULE, HWND, S_FALSE, S_OK,
};
use windows::Win32::System::Com::{IClassFactory, IClassFactory_Impl, IDispatch};
use windows::Win32::System::LibraryLoader::{
    GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
    GetModuleFileNameW, GetModuleHandleExW,
};
use windows::Win32::System::Registry::{
    HKEY, HKEY_CURRENT_USER, KEY_WRITE, REG_DWORD, REG_OPTION_NON_VOLATILE, REG_SZ, RegCloseKey,
    RegCreateKeyExW, RegDeleteTreeW, RegSetValueExW,
};
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::UI::Accessibility::AccessibleObjectFromWindow;
use windows::Win32::UI::WindowsAndMessaging::{FindWindowExW, KillTimer, SetTimer};
use windows::core::{BOOL, GUID, HRESULT, IUnknown, Interface, PCWSTR, Ref, Result, implement, w};

use super::{Addin, dispatch, is_connected};

pub const PROGID: &str = "XllBridge.Ribbon";
pub(super) const CLSID_ADDIN: GUID = GUID::from_u128(0x3C1D5B7A_6E24_4F0B_9A52_7B8D1E4C9F63);
pub(super) const CLSID_STR: &str = "{3C1D5B7A-6E24-4F0B-9A52-7B8D1E4C9F63}";
const OBJID_NATIVEOM: u32 = 0xFFFF_FFF0;

// ------------------------------------------------------------------ class factory

#[derive(Clone, Copy)]
enum Kind {
    Addin,
    Rtd,
}

#[implement(IClassFactory)]
struct Factory(Kind);

impl IClassFactory_Impl for Factory_Impl {
    fn CreateInstance(
        &self,
        outer: Ref<IUnknown>,
        iid: *const GUID,
        object: *mut *mut c_void,
    ) -> Result<()> {
        if outer.is_some() {
            return Err(CLASS_E_NOAGGREGATION.into());
        }
        let unknown: IUnknown = match self.0 {
            Kind::Addin => Addin.into(),
            Kind::Rtd => super::rtd::Rtd.into(),
        };
        // SAFETY: `iid`/`object` come straight from the COM caller.
        unsafe { unknown.query(iid, object).ok() }
    }

    fn LockServer(&self, _: BOOL) -> Result<()> {
        Ok(())
    }
}

#[unsafe(no_mangle)]
pub extern "system" fn DllGetClassObject(
    clsid: *const GUID,
    iid: *const GUID,
    object: *mut *mut c_void,
) -> HRESULT {
    catch_unwind(|| {
        if clsid.is_null() || iid.is_null() || object.is_null() {
            return E_POINTER;
        }
        // SAFETY: non-null pointers from the COM runtime.
        let kind = unsafe {
            *object = std::ptr::null_mut();
            if *clsid == CLSID_ADDIN {
                Kind::Addin
            } else if *clsid == super::rtd::CLSID_RTD {
                Kind::Rtd
            } else {
                return windows::Win32::Foundation::CLASS_E_CLASSNOTAVAILABLE;
            }
        };
        let factory: IClassFactory = Factory(kind).into();
        // SAFETY: as above.
        unsafe { factory.query(iid, object) }
    })
    .unwrap_or(E_FAIL)
}

/// This DLL is an XLL first; never ask to be unloaded.
#[unsafe(no_mangle)]
pub extern "system" fn DllCanUnloadNow() -> HRESULT {
    S_FALSE
}

#[unsafe(no_mangle)]
pub extern "system" fn DllRegisterServer() -> HRESULT {
    match catch_unwind(register) {
        Ok(Ok(())) => S_OK,
        _ => E_FAIL,
    }
}

#[unsafe(no_mangle)]
pub extern "system" fn DllUnregisterServer() -> HRESULT {
    match catch_unwind(unregister) {
        Ok(Ok(())) => S_OK,
        _ => E_FAIL,
    }
}

// ------------------------------------------------------------------ registration (HKCU only)

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(Some(0)).collect()
}

fn set_value(subkey: &str, name: Option<&str>, value: Value<'_>) -> Result<()> {
    let key_name = wide(subkey);
    let mut key = HKEY::default();
    // SAFETY: standard registry calls with NUL-terminated names that outlive them.
    unsafe {
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(key_name.as_ptr()),
            None,
            PCWSTR::null(),
            REG_OPTION_NON_VOLATILE,
            KEY_WRITE,
            None,
            &mut key,
            None,
        )
        .ok()?;
        let name_w = name.map(wide);
        let name_ptr = name_w
            .as_ref()
            .map_or(PCWSTR::null(), |n| PCWSTR(n.as_ptr()));
        let result = match value {
            Value::Str(s) => {
                let data: Vec<u8> = wide(s).iter().flat_map(|u| u.to_le_bytes()).collect();
                RegSetValueExW(key, name_ptr, None, REG_SZ, Some(&data))
            }
            Value::Dword(d) => {
                RegSetValueExW(key, name_ptr, None, REG_DWORD, Some(&d.to_le_bytes()))
            }
        };
        let _ = RegCloseKey(key);
        result.ok()
    }
}

enum Value<'a> {
    Str(&'a str),
    Dword(u32),
}

/// Full path of this DLL, exactly as the loader knows it (so Excel's LoadLibrary of it later
/// resolves to this same module instance).
fn own_path() -> Result<String> {
    let mut module = HMODULE::default();
    // SAFETY: the address of a function in this module identifies it; refcount unchanged.
    unsafe {
        GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            PCWSTR(DllCanUnloadNow as *const () as *const u16),
            &mut module,
        )?;
        let mut buf = vec![0u16; 1024];
        let n = GetModuleFileNameW(Some(module), &mut buf) as usize;
        if n == 0 {
            return Err(E_FAIL.into());
        }
        Ok(String::from_utf16_lossy(&buf[..n]))
    }
}

/// Registers the COM class and the Excel add-in entry for the current user, then arranges for
/// the add-in to be connected right away (otherwise it would only load on the next start).
pub fn register() -> Result<()> {
    let dll = own_path()?;
    let class = format!("Software\\Classes\\CLSID\\{CLSID_STR}");
    set_value(&class, None, Value::Str("xll-bridge Lua ribbon"))?;
    set_value(&format!("{class}\\InprocServer32"), None, Value::Str(&dll))?;
    set_value(
        &format!("{class}\\InprocServer32"),
        Some("ThreadingModel"),
        Value::Str("Apartment"),
    )?;
    set_value(&format!("{class}\\ProgID"), None, Value::Str(PROGID))?;
    set_value(
        &format!("Software\\Classes\\{PROGID}"),
        None,
        Value::Str("xll-bridge Lua ribbon"),
    )?;
    set_value(
        &format!("Software\\Classes\\{PROGID}\\CLSID"),
        None,
        Value::Str(CLSID_STR),
    )?;

    // The real-time data server behind LUA.STREAM.
    let rtd_class = format!("Software\\Classes\\CLSID\\{}", super::rtd::CLSID_RTD_STR);
    set_value(&rtd_class, None, Value::Str("xll-bridge Lua streams"))?;
    set_value(
        &format!("{rtd_class}\\InprocServer32"),
        None,
        Value::Str(&dll),
    )?;
    set_value(
        &format!("{rtd_class}\\InprocServer32"),
        Some("ThreadingModel"),
        Value::Str("Apartment"),
    )?;
    set_value(
        &format!("{rtd_class}\\ProgID"),
        None,
        Value::Str(super::rtd::PROGID_RTD),
    )?;
    set_value(
        &format!("Software\\Classes\\{}", super::rtd::PROGID_RTD),
        None,
        Value::Str("xll-bridge Lua streams"),
    )?;
    set_value(
        &format!("Software\\Classes\\{}\\CLSID", super::rtd::PROGID_RTD),
        None,
        Value::Str(super::rtd::CLSID_RTD_STR),
    )?;

    let addin = format!("Software\\Microsoft\\Office\\Excel\\Addins\\{PROGID}");
    set_value(&addin, Some("FriendlyName"), Value::Str("xll-bridge Lua"))?;
    set_value(
        &addin,
        Some("Description"),
        Value::Str("Lua trust indicator for xll-bridge"),
    )?;
    set_value(&addin, Some("LoadBehavior"), Value::Dword(3))?; // load at startup

    schedule_connect();
    Ok(())
}

pub fn unregister() -> Result<()> {
    for key in [
        format!("Software\\Microsoft\\Office\\Excel\\Addins\\{PROGID}"),
        format!("Software\\Classes\\CLSID\\{CLSID_STR}"),
        format!("Software\\Classes\\{PROGID}"),
        format!("Software\\Classes\\CLSID\\{}", super::rtd::CLSID_RTD_STR),
        format!("Software\\Classes\\{}", super::rtd::PROGID_RTD),
    ] {
        let k = wide(&key);
        // SAFETY: NUL-terminated key name; a missing key is not an error we care about.
        let _ = unsafe { RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(k.as_ptr())) };
    }
    Ok(())
}

// ------------------------------------------------------------------ connect without restarting

static CONNECT_TIMER: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

/// After first-time registration Excel does not know about the add-in until it restarts.
/// Give Excel a moment to finish loading this XLL, then flip `COMAddIn.Connect` ourselves.
fn schedule_connect() {
    // SAFETY: thread timer on the calling (Excel UI) thread.
    let id = unsafe { SetTimer(None, 0, 1500, Some(connect_timer)) };
    CONNECT_TIMER.store(id, std::sync::atomic::Ordering::SeqCst);
}

unsafe extern "system" fn connect_timer(_: HWND, _: u32, _: usize, _: u32) {
    let id = CONNECT_TIMER.swap(0, std::sync::atomic::Ordering::SeqCst);
    if id != 0 {
        // SAFETY: id was returned by SetTimer on this thread.
        let _ = unsafe { KillTimer(None, id) };
    }
    let _ = catch_unwind(|| {
        if !is_connected() {
            let _ = connect_now();
        }
    });
}

/// Finds Excel's `Application` object from its window handle (the classic OBJID_NATIVEOM trick).
fn excel_application(hwnd: HWND) -> Option<IDispatch> {
    // SAFETY: Win32 window lookups with static class names.
    unsafe {
        let desk = FindWindowExW(Some(hwnd), None, w!("XLDESK"), PCWSTR::null()).ok()?;
        let sheet = FindWindowExW(Some(desk), None, w!("EXCEL7"), PCWSTR::null()).ok()?;
        let mut raw: *mut c_void = std::ptr::null_mut();
        AccessibleObjectFromWindow(sheet, OBJID_NATIVEOM, &IDispatch::IID, &mut raw).ok()?;
        let window = IDispatch::from_raw(raw);
        IDispatch::try_from(&dispatch::get(&window, "Application").ok()?).ok()
    }
}

fn connect_now() -> Result<()> {
    let hwnd = super::super::excel_main_window().ok_or(E_FAIL)?;
    let app = excel_application(hwnd).ok_or(E_FAIL)?;
    let addins = IDispatch::try_from(&dispatch::get(&app, "COMAddIns")?)?;
    let item = IDispatch::try_from(&dispatch::get_with(
        &addins,
        "Item",
        &[VARIANT::from(PROGID)],
    )?)?;
    dispatch::put(&item, "Connect", VARIANT::from(true))
}
