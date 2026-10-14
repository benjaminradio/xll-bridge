//! Windows-only Excel add-in layer: exports, registration, and the ribbon status button.
//!
//! Worksheet functions (all thread-unsafe on purpose, so Lua runs on Excel's main thread and a
//! single non-`Send` Lua state per workbook is enough):
//!
//! * `LUA.CALL(name, [a1..a10])` — call a global Lua function defined by the workbook's scripts
//! * `LUA.EVAL(code, [a1..a10])` — evaluate a chunk/expression; arguments are `...`
//! * `LUA.STATUS()`              — scripts status for this workbook (none / blocked / enabled)

mod ffi;
mod oper;
mod ribbon;

use std::cell::RefCell;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::PathBuf;

use ffi::*;

use crate::runtime::{CallError, Runtime, Status, workbook_path_from_cell_filename};
use crate::trust::TrustStore;
use crate::value::Value;

thread_local! {
    static RUNTIME: RefCell<Runtime> = RefCell::new(Runtime::new(
        TrustStore::default_path().map(TrustStore::load).unwrap_or_else(TrustStore::ephemeral),
    ));
}

pub(crate) fn with_runtime<R>(f: impl FnOnce(&mut Runtime) -> R) -> R {
    RUNTIME.with(|r| f(&mut r.borrow_mut()))
}

/// Path of the workbook containing the calling cell, via `CELL("filename", caller)`.
fn caller_workbook() -> Option<PathBuf> {
    let mut caller = Xloper12::nil();
    if excel12(xlfCaller, Some(&mut caller), &mut []) != xlretSuccess {
        return None;
    }
    let what = ffi::counted_utf16("filename");
    let mut what = Xloper12::borrowed_str(&what);
    let mut out = Xloper12::nil();
    let rc = excel12(xlfCell, Some(&mut out), &mut [&mut what, &mut caller]);
    let path = (rc == xlretSuccess && out.base_type() == xltypeStr)
        .then(|| unsafe { read_counted(out.val.str_) })
        .and_then(|s| workbook_path_from_cell_filename(&s));
    free(&mut out);
    free(&mut caller);
    path
}

fn error_text(e: &CallError) -> Value {
    Value::Str(e.to_string())
}

/// Runs `f`, converting panics to a cell-sized error: unwinding into Excel would crash it.
fn guarded(f: impl FnOnce() -> Value) -> *mut Xloper12 {
    let v = catch_unwind(AssertUnwindSafe(f))
        .unwrap_or_else(|_| Value::Str("#LUA: internal error".into()));
    oper::into_result(&v)
}

fn dispatch(
    args: &[*const Xloper12],
    run: impl FnOnce(&mut Runtime, Option<&std::path::Path>, &str, &[Value]) -> Result<Value, CallError>,
) -> Value {
    // SAFETY: Excel passes valid operands (or null/missing) for every registered argument.
    let vals = unsafe { oper::collect_args(args) };
    let Some(Value::Str(first)) = vals.first().cloned() else {
        return Value::Str("#LUA: first argument must be text".into());
    };
    let wb = caller_workbook();
    let result = with_runtime(|rt| run(rt, wb.as_deref(), &first, &vals[1..]));
    result.unwrap_or_else(|e| error_text(&e))
}

/// Parameters of the `LUA.CALL`/`LUA.EVAL` entry points: the name/code plus ten arguments.
/// The registered type text is derived from this, and `udf!` builds a `[_; UDF_PARAMS]` array
/// from its parameter list, so the two cannot drift apart (they once did, and Excel crashed
/// reading an argument it never passed).
const UDF_PARAMS: usize = 11;

macro_rules! udf {
    ($name:ident, $run:expr) => {
        #[unsafe(no_mangle)]
        pub extern "system" fn $name(
            f: *const Xloper12,
            a1: *const Xloper12,
            a2: *const Xloper12,
            a3: *const Xloper12,
            a4: *const Xloper12,
            a5: *const Xloper12,
            a6: *const Xloper12,
            a7: *const Xloper12,
            a8: *const Xloper12,
            a9: *const Xloper12,
            a10: *const Xloper12,
        ) -> *mut Xloper12 {
            let all: [*const Xloper12; UDF_PARAMS] = [f, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10];
            guarded(|| dispatch(&all, $run))
        }
    };
}

udf!(LuaCall, |rt, wb, name, args| rt.call(wb, name, args));
udf!(LuaEval, |rt, wb, code, args| rt.eval(wb, code, args));

#[unsafe(no_mangle)]
pub extern "system" fn LuaStatus() -> *mut Xloper12 {
    guarded(|| {
        let Some(wb) = caller_workbook() else {
            return Value::Str("unsaved workbook".into());
        };
        let s = with_runtime(|rt| match rt.status(&wb) {
            Status::NoScripts => "no Lua scripts".to_owned(),
            Status::Blocked { names } => {
                format!(
                    "{} script(s) BLOCKED - use the Lua ribbon button",
                    names.len()
                )
            }
            Status::Active { .. } => "Lua enabled".to_owned(),
            Status::Failed { message } => format!("Lua failed to load: {message}"),
        });
        Value::Str(s)
    })
}

// ------------------------------------------------------------------ add-in lifecycle

struct FnDef {
    proc_name: &'static str,
    /// Number of worksheet arguments the exported function takes.
    params: usize,
    excel_name: &'static str,
    arg_text: &'static str,
    help: &'static str,
}

const FUNCTIONS: [FnDef; 3] = [
    FnDef {
        proc_name: "LuaCall",
        params: UDF_PARAMS,
        excel_name: "LUA.CALL",
        arg_text: "function,arg1,arg2,arg3,arg4,arg5,arg6,arg7,arg8,arg9,arg10",
        help: "Calls a global Lua function defined by this workbook's embedded scripts",
    },
    FnDef {
        proc_name: "LuaEval",
        params: UDF_PARAMS,
        excel_name: "LUA.EVAL",
        arg_text: "code,arg1,arg2,arg3,arg4,arg5,arg6,arg7,arg8,arg9,arg10",
        help: "Evaluates a Lua expression or chunk; arguments are available as ...",
    },
    FnDef {
        proc_name: "LuaStatus",
        params: 0,
        excel_name: "LUA.STATUS",
        arg_text: "",
        help: "Shows whether this workbook's Lua scripts are enabled",
    },
];

fn register_all() -> bool {
    let mut dll = Xloper12::nil();
    if excel12(xlGetName, Some(&mut dll), &mut []) != xlretSuccess {
        return false;
    }
    for f in &FUNCTIONS {
        // Type text is the return type followed by one code per argument; all are `Q` (XLOPER12).
        let type_text = "Q".repeat(1 + f.params);
        let bufs: Vec<Vec<u16>> = [
            f.proc_name,
            type_text.as_str(),
            f.excel_name,
            f.arg_text,
            "Lua",
            f.help,
        ]
        .iter()
        .map(|s| counted_utf16(s))
        .collect();
        let mut s: Vec<Xloper12> = bufs.iter().map(|b| Xloper12::borrowed_str(b)).collect();
        // xlfRegister(dll, procedure, type, name, args, macroType=1, category, shortcut, helpTopic, help)
        let mut macro_type = Xloper12::int(1);
        let (mut shortcut, mut topic) = (Xloper12::missing(), Xloper12::missing());
        let [proc_, ty, name, args, cat, help] = &mut s[..] else {
            unreachable!()
        };
        let mut res = Xloper12::nil();
        let rc = excel12(
            xlfRegister,
            Some(&mut res),
            &mut [
                &mut dll,
                proc_,
                ty,
                name,
                args,
                &mut macro_type,
                cat,
                &mut shortcut,
                &mut topic,
                help,
            ],
        );
        free(&mut res);
        if rc != xlretSuccess {
            free(&mut dll);
            return false;
        }
    }
    free(&mut dll);
    true
}

/// Records panics and access violations in `addin.log`. Excel's own crash dialog says nothing
/// about the cause; this does. First-chance exceptions Excel handles itself are logged too, so
/// the count is capped and a line is not proof of a crash on its own.
fn install_diagnostics() {
    use std::sync::Once;
    use std::sync::atomic::{AtomicU32, Ordering};
    use windows::Win32::System::Diagnostics::Debug::{
        AddVectoredExceptionHandler, EXCEPTION_POINTERS,
    };
    use windows::Win32::System::LibraryLoader::{
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        GetModuleFileNameW, GetModuleHandleExW,
    };
    use windows::core::PCWSTR;

    unsafe extern "system" fn veh(info: *mut EXCEPTION_POINTERS) -> i32 {
        static SEEN: AtomicU32 = AtomicU32::new(0);
        // SAFETY: the OS passes valid exception pointers.
        let (code, addr) = unsafe {
            let rec = (*info).ExceptionRecord;
            (
                (*rec).ExceptionCode.0 as u32,
                (*rec).ExceptionAddress as usize,
            )
        };
        // Access violation, stack overflow, fast-fail.
        if matches!(code, 0xC000_0005 | 0xC000_00FD | 0xC000_0409)
            && SEEN.fetch_add(1, Ordering::Relaxed) < 20
        {
            let mut place = String::from("unknown module");
            // SAFETY: address-based module lookup; buffer sized for a long path.
            unsafe {
                let mut module = windows::Win32::Foundation::HMODULE::default();
                if GetModuleHandleExW(
                    GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS
                        | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                    PCWSTR(addr as *const u16),
                    &mut module,
                )
                .is_ok()
                {
                    let mut buf = vec![0u16; 520];
                    let n = GetModuleFileNameW(Some(module), &mut buf) as usize;
                    place = format!(
                        "{}+{:#x}",
                        String::from_utf16_lossy(&buf[..n]),
                        addr - module.0 as usize
                    );
                }
            }
            crate::log::line(&format!("exception {code:#010x} at {addr:#x} in {place}"));
        }
        0 // EXCEPTION_CONTINUE_SEARCH: observe only
    }

    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        std::panic::set_hook(Box::new(|info| crate::log::line(&format!("panic: {info}"))));
        // SAFETY: registers a process-wide observer; it never alters the exception flow.
        unsafe { AddVectoredExceptionHandler(1, Some(veh)) };
    });
}

/// Main Excel window, for the COM bootstrap.
pub(crate) fn excel_main_window() -> Option<windows::Win32::Foundation::HWND> {
    let mut out = Xloper12::nil();
    if excel12(xlGetHwnd, Some(&mut out), &mut []) != xlretSuccess {
        return None;
    }
    // Excel returns the HWND as a 32-bit integer.
    let h = unsafe { out.val.w } as u32 as usize;
    Some(windows::Win32::Foundation::HWND(
        h as *mut core::ffi::c_void,
    ))
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoOpen() -> i32 {
    catch_unwind(|| {
        install_diagnostics();
        crate::log::note(&format!("xlAutoOpen v{}", env!("CARGO_PKG_VERSION")));
        let ok = register_all();
        // The ribbon button is optional: UDFs must keep working if COM registration fails.
        match ribbon::register() {
            Ok(()) => crate::log::note("ribbon: registered"),
            Err(e) => crate::log::line(&format!("ribbon: registration failed: {e}")),
        }
        ok
    })
    .unwrap_or(false) as i32
}

/// Excel exiting, or the add-in being unloaded. The ribbon registration is left in place.
#[unsafe(no_mangle)]
pub extern "system" fn xlAutoClose() -> i32 {
    1
}

/// The user removed the add-in from the Add-ins dialog: take the ribbon registration with it.
#[unsafe(no_mangle)]
pub extern "system" fn xlAutoRemove() -> i32 {
    let _ = catch_unwind(ribbon::unregister);
    1
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoFree12(p: *mut Xloper12) {
    // SAFETY: Excel only hands back pointers we returned with xlbitDLLFree.
    let _ = catch_unwind(AssertUnwindSafe(|| unsafe { oper::free_result(p) }));
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAddInManagerInfo12(action: *const Xloper12) -> *mut Xloper12 {
    // SAFETY: Excel passes a valid operand.
    let wants_name = matches!(unsafe { oper::to_value(action) }, Value::Num(n) if n == 1.0);
    oper::into_result(&if wants_name {
        Value::Str("xll-bridge (Lua)".into())
    } else {
        Value::Err(crate::value::ErrorCode::Value)
    })
}
