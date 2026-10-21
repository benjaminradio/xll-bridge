//! Late-bound (IDispatch) helpers for talking to Excel's object model without a type library.

use std::ptr::null_mut;

use windows::Win32::Foundation::DISP_E_EXCEPTION;
use windows::Win32::System::Com::{
    DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS,
    EXCEPINFO, IDispatch,
};
use windows::Win32::System::Variant::VARIANT;
use windows::core::{GUID, HRESULT, PCWSTR, Result};

const LOCALE_USER_DEFAULT: u32 = 0x0400;
const DISPID_PROPERTYPUT: i32 = -3;

fn dispid(obj: &IDispatch, name: &str) -> Result<i32> {
    let wide: Vec<u16> = name.encode_utf16().chain(Some(0)).collect();
    let names = [PCWSTR(wide.as_ptr())];
    let mut id = 0i32;
    // SAFETY: `names` and `id` outlive the call; one name requested.
    unsafe {
        obj.GetIDsOfNames(
            &GUID::zeroed(),
            names.as_ptr(),
            1,
            LOCALE_USER_DEFAULT,
            &mut id,
        )?
    };
    Ok(id)
}

fn invoke(obj: &IDispatch, name: &str, flags: DISPATCH_FLAGS, args: &[VARIANT]) -> Result<VARIANT> {
    let id = dispid(obj, name)?;
    // IDispatch takes arguments right-to-left.
    let mut rev: Vec<VARIANT> = args.iter().rev().cloned().collect();
    let mut named = [DISPID_PROPERTYPUT];
    let is_put = flags == DISPATCH_PROPERTYPUT;
    let params = DISPPARAMS {
        rgvarg: if rev.is_empty() {
            null_mut()
        } else {
            rev.as_mut_ptr()
        },
        rgdispidNamedArgs: if is_put {
            named.as_mut_ptr()
        } else {
            null_mut()
        },
        cArgs: rev.len() as u32,
        cNamedArgs: u32::from(is_put),
    };
    let mut result = VARIANT::default();
    // SAFETY: an all-zero EXCEPINFO is the documented empty state.
    let mut excep: EXCEPINFO = unsafe { std::mem::zeroed() };
    // SAFETY: `params` points at live argument storage for the duration of the call.
    let outcome = unsafe {
        obj.Invoke(
            id,
            &GUID::zeroed(),
            LOCALE_USER_DEFAULT,
            flags,
            &params,
            Some(&mut result),
            Some(&mut excep),
            None,
        )
    };
    if let Err(e) = outcome {
        // "Exception occurred" says nothing; the callee's own description does.
        let mut text = String::new();
        if e.code() == DISP_E_EXCEPTION {
            // SAFETY: the callee filled these BSTRs; each is read once and freed once.
            unsafe {
                let source = excep.bstrSource.to_string();
                let desc = excep.bstrDescription.to_string();
                std::mem::ManuallyDrop::drop(&mut excep.bstrSource);
                std::mem::ManuallyDrop::drop(&mut excep.bstrDescription);
                std::mem::ManuallyDrop::drop(&mut excep.bstrHelpFile);
                text = match (source.is_empty(), desc.is_empty()) {
                    (_, true) => String::new(),
                    (true, false) => desc,
                    (false, false) => format!("{source}: {desc}"),
                };
            }
        }
        if !text.is_empty() {
            let code = if excep.scode != 0 {
                excep.scode
            } else {
                e.code().0
            };
            return Err(windows::core::Error::new(HRESULT(code), text));
        }
        return Err(e);
    }
    Ok(result)
}

pub fn get(obj: &IDispatch, name: &str) -> Result<VARIANT> {
    invoke(obj, name, DISPATCH_PROPERTYGET, &[])
}

pub fn get_with(obj: &IDispatch, name: &str, args: &[VARIANT]) -> Result<VARIANT> {
    // Parameterised properties (e.g. `COMAddIns.Item("id")`) are usually callable as methods too.
    invoke(obj, name, DISPATCH_METHOD | DISPATCH_PROPERTYGET, args)
}

pub fn put(obj: &IDispatch, name: &str, value: VARIANT) -> Result<()> {
    invoke(obj, name, DISPATCH_PROPERTYPUT, &[value]).map(drop)
}

pub fn call(obj: &IDispatch, name: &str, args: &[VARIANT]) -> Result<VARIANT> {
    invoke(obj, name, DISPATCH_METHOD, args)
}
