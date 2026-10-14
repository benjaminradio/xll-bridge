#![allow(non_upper_case_globals)] // xltype* names follow XLCALL.H
//! XLOPER12 <-> [`Value`] conversion and ownership of values handed back to Excel.
//!
//! Results returned from a UDF are boxed and flagged `xlbitDLLFree`; Excel then calls
//! `xlAutoFree12`, which lands in [`free_result`].

use super::ffi::*;
use crate::value::{Array, ErrorCode, Value};

/// Converts an argument operand. References are coerced to values first.
///
/// # Safety
/// `p` must be null or point to a valid XLOPER12 supplied by Excel.
pub unsafe fn to_value(p: *const Xloper12) -> Value {
    if p.is_null() {
        return Value::Empty;
    }
    let x = unsafe { &*p };
    match x.base_type() {
        xltypeNum => Value::Num(unsafe { x.val.num }),
        xltypeInt => Value::Num(unsafe { x.val.w } as f64),
        xltypeStr => Value::Str(unsafe { read_counted(x.val.str_) }),
        xltypeBool => Value::Bool(unsafe { x.val.xbool } != 0),
        xltypeErr => Value::Err(ErrorCode::from_xl_code(unsafe { x.val.err }).unwrap_or(ErrorCode::Value)),
        xltypeMulti => {
            let a = unsafe { x.val.array };
            let (rows, cols) = (a.rows.max(0) as usize, a.columns.max(0) as usize);
            let data = (0..rows * cols).map(|i| unsafe { to_value(a.lparray.add(i)) }).collect();
            Value::Array(Array::new(rows, cols, data))
        }
        xltypeSRef | xltypeRef => {
            let mut src = *x;
            let mut out = Xloper12::nil();
            if excel12(xlCoerce, Some(&mut out), &mut [&mut src]) != xlretSuccess {
                return Value::Err(ErrorCode::Value);
            }
            let v = unsafe { to_value(&out) };
            free(&mut out);
            v
        }
        // xltypeMissing, xltypeNil, anything exotic
        _ => Value::Empty,
    }
}

/// Number of non-missing leading arguments (trailing omitted ones are dropped).
///
/// # Safety
/// Every non-null pointer must be a valid XLOPER12.
pub unsafe fn collect_args(ptrs: &[*const Xloper12]) -> Vec<Value> {
    let n = ptrs
        .iter()
        .rposition(|&p| !p.is_null() && unsafe { (*p).base_type() } != xltypeMissing)
        .map_or(0, |i| i + 1);
    ptrs[..n].iter().map(|&p| unsafe { to_value(p) }).collect()
}

fn leak_str(s: &str) -> *mut u16 {
    Box::into_raw(counted_utf16(s).into_boxed_slice()) as *mut u16
}

fn fill(v: &Value) -> Xloper12 {
    let mut out = Xloper12::nil();
    match v {
        Value::Empty => {}
        Value::Num(n) => {
            out.val.num = *n;
            out.xltype = xltypeNum;
        }
        Value::Bool(b) => {
            out.val.xbool = *b as i32;
            out.xltype = xltypeBool;
        }
        Value::Str(s) => {
            out.val.str_ = leak_str(s);
            out.xltype = xltypeStr;
        }
        Value::Err(e) => {
            out.val.err = e.xl_code();
            out.xltype = xltypeErr;
        }
        Value::Array(a) => {
            let cells: Vec<Xloper12> = a
                .data
                .iter()
                .map(|c| match c {
                    // Nested arrays aren't representable in a cell.
                    Value::Array(_) => fill(&Value::Err(ErrorCode::Value)),
                    other => fill(other),
                })
                .collect();
            let ptr = Box::into_raw(cells.into_boxed_slice()) as *mut Xloper12;
            out.val.array = MultiArray { lparray: ptr, rows: a.rows as i32, columns: a.cols as i32 };
            out.xltype = xltypeMulti;
        }
    }
    out
}

/// Boxes `v` for return to Excel. Freed by [`free_result`] via `xlAutoFree12`.
pub fn into_result(v: &Value) -> *mut Xloper12 {
    let mut x = fill(v);
    x.xltype |= xlbitDLLFree;
    Box::into_raw(Box::new(x))
}

/// Frees only what [`into_result`]/[`fill`] allocated.
///
/// # Safety
/// `p` must come from [`into_result`].
pub unsafe fn free_result(p: *mut Xloper12) {
    if p.is_null() {
        return;
    }
    let x = unsafe { Box::from_raw(p) };
    unsafe { free_inner(&x) };
}

unsafe fn free_inner(x: &Xloper12) {
    match x.base_type() {
        xltypeStr => {
            let p = unsafe { x.val.str_ };
            if !p.is_null() {
                let len = unsafe { *p } as usize + 1;
                drop(unsafe { Box::from_raw(std::ptr::slice_from_raw_parts_mut(p, len)) });
            }
        }
        xltypeMulti => {
            let a = unsafe { x.val.array };
            let n = (a.rows * a.columns).max(0) as usize;
            let cells = unsafe { Box::from_raw(std::ptr::slice_from_raw_parts_mut(a.lparray, n)) };
            for c in cells.iter() {
                unsafe { free_inner(c) };
            }
        }
        _ => {}
    }
}
