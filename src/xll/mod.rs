//! Windows-only Excel add-in layer: exports, registration, and the trust banner.
//!
//! Worksheet functions (all thread-unsafe on purpose, so Lua runs on Excel's main thread and a
//! single non-`Send` Lua state per workbook is enough):
//!
//! * `LUA.CALL(name, [a1..a10])` — call a global Lua function defined by the workbook's scripts
//! * `LUA.EVAL(code, [a1..a10])` — evaluate a chunk/expression; arguments are `...`
//! * `LUA.STATUS()`              — scripts status for this workbook (none / blocked / enabled)

mod banner;
mod ffi;
mod oper;

use std::cell::RefCell;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::PathBuf;

use ffi::*;

use crate::runtime::{Book, CallError, Runtime, workbook_path_from_cell_filename};
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
    let result = with_runtime(|rt| {
        let r = run(rt, wb.as_deref(), &first, &vals[1..]);
        // Raise the banner from the same call that was refused.
        if let (Err(CallError::Blocked), Some(path)) = (&r, wb.as_deref())
            && let Some((_, n)) = rt.is_blocked(path)
        {
            banner::show(path, n);
        }
        r
    });
    result.unwrap_or_else(|e| error_text(&e))
}

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
            guarded(|| dispatch(&[f, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10], $run))
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
        let s = with_runtime(|rt| match rt.state(&wb) {
            Book::NoScripts => "no Lua scripts".to_owned(),
            Book::Blocked { scripts, .. } => {
                format!("{scripts} script(s) DISABLED - click Enable Lua")
            }
            Book::Active { .. } => "Lua enabled".to_owned(),
            Book::Failed { message, .. } => format!("Lua failed to load: {message}"),
        });
        Value::Str(s)
    })
}

// ------------------------------------------------------------------ add-in lifecycle

struct FnDef {
    proc_name: &'static str,
    type_text: &'static str,
    excel_name: &'static str,
    arg_text: &'static str,
    help: &'static str,
}

const FUNCTIONS: [FnDef; 3] = [
    FnDef {
        proc_name: "LuaCall",
        type_text: "QQQQQQQQQQQ",
        excel_name: "LUA.CALL",
        arg_text: "function,arg1,arg2,arg3,arg4,arg5,arg6,arg7,arg8,arg9,arg10",
        help: "Calls a global Lua function defined by this workbook's embedded scripts",
    },
    FnDef {
        proc_name: "LuaEval",
        type_text: "QQQQQQQQQQQ",
        excel_name: "LUA.EVAL",
        arg_text: "code,arg1,arg2,arg3,arg4,arg5,arg6,arg7,arg8,arg9,arg10",
        help: "Evaluates a Lua expression or chunk; arguments are available as ...",
    },
    FnDef {
        proc_name: "LuaStatus",
        type_text: "Q",
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
        let bufs: Vec<Vec<u16>> = [
            f.proc_name,
            f.type_text,
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

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoOpen() -> i32 {
    catch_unwind(register_all).unwrap_or(false) as i32
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoClose() -> i32 {
    let _ = catch_unwind(banner::dismiss);
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
