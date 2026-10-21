//! The Excel side of `xl.range`: executes [`crate::xlapi::Op`]s through the object model.
//!
//! Everything here runs on Excel's UI thread. Lua code on other threads reaches it through the
//! [`crate::xlapi::Bridge`]; a hidden message-only window is posted a message to run the queue
//! promptly (the 200 ms timer pumps it too, as a safety net).

use std::cell::Cell;
use std::path::Path;
use std::sync::Arc;

use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::Com::{IDispatch, SAFEARRAY, SAFEARRAYBOUND};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Ole::{
    SafeArrayAccessData, SafeArrayCreate, SafeArrayGetDim, SafeArrayGetLBound, SafeArrayGetUBound,
    SafeArrayUnaccessData,
};
use windows::Win32::System::Variant::{
    VARENUM, VARIANT, VT_ARRAY, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_EMPTY, VT_ERROR, VT_I1, VT_I2,
    VT_I4, VT_I8, VT_NULL, VT_R4, VT_R8, VT_UI1, VT_UI2, VT_UI4, VT_VARIANT,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, HWND_MESSAGE, PostMessageW, RegisterClassW, WINDOW_EX_STYLE,
    WINDOW_STYLE, WM_APP, WNDCLASSW,
};
use windows::core::{BSTR, w};

use super::dispatch;
use crate::value::{Array, ErrorCode, Value};
use crate::xlapi::{self, Addr, Bridge, Op, RangeInfo, Reply, XlError, quote_sheet, split_address};

// ------------------------------------------------------------------ Value <-> VARIANT

/// Excel reports worksheet errors as `VT_ERROR` with `scode = 0x800A0000 + 2000 + code`.
const XL_ERROR_BASE: i32 = 0x800A_07D0_u32 as i32;

fn error_scode(e: ErrorCode) -> i32 {
    XL_ERROR_BASE.wrapping_add(e.xl_code())
}

fn error_from_scode(scode: i32) -> ErrorCode {
    ErrorCode::from_xl_code(scode.wrapping_sub(XL_ERROR_BASE)).unwrap_or(ErrorCode::Value)
}

pub(super) fn variant_of(v: &Value) -> VARIANT {
    match v {
        Value::Empty => VARIANT::default(),
        Value::Num(n) => VARIANT::from(*n),
        Value::Bool(b) => VARIANT::from(*b),
        Value::Str(s) => VARIANT::from(s.as_str()),
        Value::Err(e) => {
            let mut out = VARIANT::default();
            // SAFETY: writing the discriminant and the matching union member of a fresh VARIANT.
            unsafe {
                let inner = &mut *out.Anonymous.Anonymous;
                inner.vt = VT_ERROR;
                inner.Anonymous.scode = error_scode(*e);
            }
            out
        }
        // Nested arrays make no sense inside a cell.
        Value::Array(_) => VARIANT::from("#VALUE!"),
    }
}

fn value_of(v: &VARIANT) -> Value {
    let vt = v.vt();
    // SAFETY: each arm reads the union member that `vt` says is active.
    unsafe {
        let a = &*v.Anonymous.Anonymous;
        match vt {
            VT_EMPTY | VT_NULL => Value::Empty,
            VT_R8 => Value::Num(a.Anonymous.dblVal),
            VT_R4 => Value::Num(f64::from(a.Anonymous.fltVal)),
            VT_DATE => Value::Num(a.Anonymous.date),
            VT_I4 => Value::Num(f64::from(a.Anonymous.lVal)),
            VT_I2 => Value::Num(f64::from(a.Anonymous.iVal)),
            VT_I1 => Value::Num(f64::from(a.Anonymous.cVal)),
            VT_UI1 => Value::Num(f64::from(a.Anonymous.bVal)),
            VT_UI2 => Value::Num(f64::from(a.Anonymous.uiVal)),
            VT_UI4 => Value::Num(f64::from(a.Anonymous.ulVal)),
            VT_I8 => Value::Num(a.Anonymous.llVal as f64),
            VT_CY => Value::Num(a.Anonymous.cyVal.int64 as f64 / 10_000.0),
            VT_BOOL => Value::Bool(a.Anonymous.boolVal.as_bool()),
            VT_BSTR => Value::Str(a.Anonymous.bstrVal.to_string()),
            VT_ERROR => Value::Err(error_from_scode(a.Anonymous.scode)),
            _ => Value::Str(v.to_string()),
        }
    }
}

const VT_ARRAY_OF_VARIANT: VARENUM = VARENUM(VT_ARRAY.0 | VT_VARIANT.0);

/// A 2-D `SAFEARRAY(VARIANT)` (what `Range.Value2` returns for several cells; 1-based, column
/// major) as an [`Array`]; a plain VARIANT as a scalar.
pub(super) fn read_variant(v: &VARIANT) -> Result<Value, XlError> {
    if v.vt() != VT_ARRAY_OF_VARIANT {
        return Ok(value_of(v));
    }
    // SAFETY: vt says `parray` is the active member; the array is owned by `v` and not modified.
    unsafe {
        let sa: *mut SAFEARRAY = v.Anonymous.Anonymous.Anonymous.parray;
        if sa.is_null() || SafeArrayGetDim(sa) != 2 {
            return Err(XlError::new("Excel returned an unexpected array shape"));
        }
        let bound = |dim| -> Result<usize, XlError> {
            let lo = SafeArrayGetLBound(sa, dim).map_err(com)?;
            let hi = SafeArrayGetUBound(sa, dim).map_err(com)?;
            Ok((hi - lo + 1).max(0) as usize)
        };
        let (rows, cols) = (bound(1)?, bound(2)?);
        let mut data_ptr = std::ptr::null_mut();
        SafeArrayAccessData(sa, &mut data_ptr).map_err(com)?;
        let items = data_ptr as *const VARIANT;
        let mut data = Vec::with_capacity(rows * cols);
        for r in 0..rows {
            for c in 0..cols {
                data.push(value_of(&*items.add(r + c * rows)));
            }
        }
        let _ = SafeArrayUnaccessData(sa);
        Ok(Value::Array(Array::new(rows, cols, data)))
    }
}

/// An [`Array`] as a 2-D `SAFEARRAY(VARIANT)` for `Range.Value2 = ...`; a scalar as a VARIANT.
pub(super) fn write_variant(v: &Value) -> Result<VARIANT, XlError> {
    let Value::Array(a) = v else {
        return Ok(variant_of(v));
    };
    if a.rows == 0 || a.cols == 0 {
        return Err(XlError::new("cannot write an empty table"));
    }
    let bounds = [
        SAFEARRAYBOUND {
            cElements: a.rows as u32,
            lLbound: 1,
        },
        SAFEARRAYBOUND {
            cElements: a.cols as u32,
            lLbound: 1,
        },
    ];
    // SAFETY: the array is created with `rows * cols` zeroed VARIANT slots, which are filled by
    // moving freshly built VARIANTs in; ownership then passes to the returned VARIANT.
    unsafe {
        let sa = SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr());
        if sa.is_null() {
            return Err(XlError::new("out of memory preparing cells for Excel"));
        }
        let mut data_ptr = std::ptr::null_mut();
        if let Err(e) = SafeArrayAccessData(sa, &mut data_ptr) {
            let _ = windows::Win32::System::Ole::SafeArrayDestroy(sa);
            return Err(com(e));
        }
        let items = data_ptr as *mut VARIANT;
        for r in 0..a.rows {
            for c in 0..a.cols {
                std::ptr::write(
                    items.add(r + c * a.rows),
                    variant_of(&a.data[r * a.cols + c]),
                );
            }
        }
        let _ = SafeArrayUnaccessData(sa);
        let mut out = VARIANT::default();
        let inner = &mut *out.Anonymous.Anonymous;
        inner.vt = VT_ARRAY_OF_VARIANT;
        inner.Anonymous.parray = sa;
        Ok(out)
    }
}

// ------------------------------------------------------------------ COM errors

const RPC_E_CALL_REJECTED: u32 = 0x8001_0001;
const RPC_E_SERVERCALL_RETRYLATER: u32 = 0x8001_010A;
const VBA_E_IGNORE: u32 = 0x800A_C472;
const DISP_E_BADINDEX: u32 = 0x8002_000B;
const E_EXCEL_BAD_REFERENCE: u32 = 0x800A_03EC;

fn com(e: windows::core::Error) -> XlError {
    let code = e.code().0 as u32;
    let text = e.message();
    if matches!(
        code,
        RPC_E_CALL_REJECTED | RPC_E_SERVERCALL_RETRYLATER | VBA_E_IGNORE
    ) {
        XlError::busy("Excel is busy (editing a cell or showing a dialog)")
    } else if text.is_empty() {
        XlError::new(format!("Excel error 0x{code:08X}"))
    } else {
        XlError::new(format!("Excel: {} (0x{code:08X})", text.trim()))
    }
}

fn dispatch_of(v: windows::core::Result<VARIANT>) -> Result<IDispatch, XlError> {
    v.and_then(|v| IDispatch::try_from(&v)).map_err(com)
}

fn string_of(v: windows::core::Result<VARIANT>) -> Result<String, XlError> {
    v.and_then(|v| BSTR::try_from(&v))
        .map(|b| b.to_string())
        .map_err(com)
}

fn count_of(obj: &IDispatch, prop: &str) -> Result<i32, XlError> {
    dispatch::get(obj, prop)
        .and_then(|v| i32::try_from(&v))
        .map_err(com)
}

// ------------------------------------------------------------------ resolving ranges

fn book_name_of(path: &Path) -> Option<String> {
    super::workbook_name_by_path(path)
}

pub(super) fn open_book(app: &IDispatch, name: &str) -> Result<IDispatch, XlError> {
    let books = dispatch_of(dispatch::get(app, "Workbooks"))?;
    dispatch_of(dispatch::get_with(&books, "Item", &[VARIANT::from(name)])).map_err(|e| {
        if e.retryable {
            e
        } else {
            XlError::new(format!("no open workbook named '{name}'"))
        }
    })
}

struct Resolved {
    range: IDispatch,
    /// Name of the workbook the range lives in, and whether the caller named it explicitly.
    book: String,
    explicit_book: bool,
}

fn resolve(app: &IDispatch, workbook: &Path, address: &str) -> Result<Resolved, XlError> {
    let addr: Addr = split_address(address).map_err(XlError::new)?;
    let default_book = book_name_of(workbook).ok_or_else(|| {
        XlError::new(
            "this workbook is not open in Excel under a known name yet; try again in a moment",
        )
    })?;
    let (book, explicit_book) = match addr.book {
        Some(b) if !b.eq_ignore_ascii_case(&default_book) => (b, true),
        _ => (default_book, false),
    };
    let wb = open_book(app, &book)?;
    let sheet = match &addr.sheet {
        Some(name) => {
            let sheets = dispatch_of(dispatch::get(&wb, "Worksheets"))?;
            dispatch_of(dispatch::get_with(
                &sheets,
                "Item",
                &[VARIANT::from(name.as_str())],
            ))
            .map_err(|e| {
                if e.retryable {
                    e
                } else {
                    XlError::new(format!("no sheet named '{name}' in {book}"))
                }
            })?
        }
        None => dispatch_of(dispatch::get(&wb, "ActiveSheet")).map_err(|_| {
            XlError::new("the active sheet is not a worksheet; name the sheet in the address")
        })?,
    };
    // `Range` accepts both cell references and defined names.
    let range = dispatch::get_with(&sheet, "Range", &[VARIANT::from(addr.cells.as_str())])
        .and_then(|v| IDispatch::try_from(&v))
        .map_err(|e| {
            if e.code().0 as u32 == E_EXCEL_BAD_REFERENCE || e.code().0 as u32 == DISP_E_BADINDEX {
                XlError::new(format!("'{}' is not a valid range or name", addr.cells))
            } else {
                com(e)
            }
        })?;
    Ok(Resolved {
        range,
        book,
        explicit_book,
    })
}

fn describe(range: &IDispatch, book: &str, explicit_book: bool) -> Result<RangeInfo, XlError> {
    let areas = dispatch_of(dispatch::get(range, "Areas"))?;
    if count_of(&areas, "Count")? != 1 {
        return Err(XlError::new(
            "ranges made of several areas are not supported",
        ));
    }
    let cells = string_of(dispatch::get(range, "Address"))?;
    let sheet = dispatch_of(dispatch::get(range, "Worksheet"))?;
    let sheet_name = string_of(dispatch::get(&sheet, "Name"))?;
    let rows = count_of(&dispatch_of(dispatch::get(range, "Rows"))?, "Count")?;
    let cols = count_of(&dispatch_of(dispatch::get(range, "Columns"))?, "Count")?;
    let address = if explicit_book {
        format!(
            "'[{}]{}'!{cells}",
            book.replace('\'', "''"),
            sheet_name.replace('\'', "''")
        )
    } else {
        format!("{}!{cells}", quote_sheet(&sheet_name))
    };
    Ok(RangeInfo {
        address,
        rows: rows.max(0) as u32,
        cols: cols.max(0) as u32,
    })
}

fn top_left(range: &IDispatch, rows: usize, cols: usize) -> Result<IDispatch, XlError> {
    let first = dispatch_of(dispatch::get_with(
        range,
        "Cells",
        &[VARIANT::from(1i32), VARIANT::from(1i32)],
    ))?;
    dispatch_of(dispatch::get_with(
        &first,
        "Resize",
        &[VARIANT::from(rows as i32), VARIANT::from(cols as i32)],
    ))
}

// ------------------------------------------------------------------ operations

/// Runs one operation against the attached Excel. Call only on the UI thread.
pub(super) fn execute(op: Op) -> Result<Reply, XlError> {
    let app = super::app().ok_or_else(|| XlError::new("Excel is not attached yet"))?;
    match op {
        Op::Info { workbook, address } => {
            let r = resolve(&app, &workbook, &address)?;
            Ok(Reply::Info(describe(&r.range, &r.book, r.explicit_book)?))
        }
        Op::Read {
            workbook,
            address,
            formulas,
        } => {
            let r = resolve(&app, &workbook, &address)?;
            let prop = if formulas { "Formula" } else { "Value2" };
            let v = dispatch::get(&r.range, prop).map_err(com)?;
            let value = read_variant(&v)?;
            // Callers always get a table of rows, even for one cell.
            Ok(Reply::Data(match value {
                Value::Array(_) => value,
                scalar => Value::Array(Array::new(1, 1, vec![scalar])),
            }))
        }
        Op::Write {
            workbook,
            address,
            formulas,
            data,
        } => {
            let r = resolve(&app, &workbook, &address)?;
            let prop = if formulas { "Formula" } else { "Value2" };
            let target = match &data {
                Value::Array(a) => top_left(&r.range, a.rows, a.cols)?,
                _ => r.range,
            };
            dispatch::put(&target, prop, write_variant(&data)?).map_err(com)?;
            Ok(Reply::Done)
        }
        Op::Clear { workbook, address } => {
            let r = resolve(&app, &workbook, &address)?;
            dispatch::call(&r.range, "ClearContents", &[]).map_err(com)?;
            Ok(Reply::Done)
        }
        Op::Sheets { workbook } => {
            let book = book_name_of(&workbook)
                .ok_or_else(|| XlError::new("this workbook is not open in Excel"))?;
            let sheets = dispatch_of(dispatch::get(&open_book(&app, &book)?, "Worksheets"))?;
            let n = count_of(&sheets, "Count")?;
            let mut names = Vec::new();
            for i in 1..=n {
                let ws = dispatch_of(dispatch::get_with(&sheets, "Item", &[VARIANT::from(i)]))?;
                names.push(string_of(dispatch::get(&ws, "Name"))?);
            }
            Ok(Reply::Names(names))
        }
        Op::UsedRange { workbook, sheet } => {
            let book = book_name_of(&workbook)
                .ok_or_else(|| XlError::new("this workbook is not open in Excel"))?;
            let wb = open_book(&app, &book)?;
            let sheets = dispatch_of(dispatch::get(&wb, "Worksheets"))?;
            let ws = dispatch_of(dispatch::get_with(
                &sheets,
                "Item",
                &[VARIANT::from(sheet.as_str())],
            ))
            .map_err(|_| XlError::new(format!("no sheet named '{sheet}'")))?;
            let used = dispatch_of(dispatch::get(&ws, "UsedRange"))?;
            Ok(Reply::Info(describe(&used, &book, false)?))
        }
        Op::Selection { workbook } => {
            let default_book = book_name_of(&workbook).unwrap_or_default();
            let sel = dispatch_of(dispatch::get(&app, "Selection"))
                .map_err(|_| XlError::new("the selection is not a range of cells"))?;
            let sheet = dispatch_of(dispatch::get(&sel, "Worksheet"))
                .map_err(|_| XlError::new("the selection is not a range of cells"))?;
            let parent = dispatch_of(dispatch::get(&sheet, "Parent"))?;
            let book = string_of(dispatch::get(&parent, "Name"))?;
            let explicit = !book.eq_ignore_ascii_case(&default_book);
            Ok(Reply::Info(describe(&sel, &book, explicit)?))
        }
    }
}

// ------------------------------------------------------------------ waking the UI thread

const WAKE_CLASS: windows::core::PCWSTR = w!("XllBridgeWake");

thread_local! {
    static WAKE_WINDOW: Cell<isize> = const { Cell::new(0) };
}

unsafe extern "system" fn wake_proc(hwnd: HWND, msg: u32, wp: WPARAM, lp: LPARAM) -> LRESULT {
    if msg == WM_APP {
        let _ = std::panic::catch_unwind(xlapi::pump_global);
        return LRESULT(0);
    }
    // SAFETY: default handling of a message we received.
    unsafe { DefWindowProcW(hwnd, msg, wp, lp) }
}

/// Starts serving worker-thread requests. Must be called on the UI thread, once Excel is attached.
pub(super) fn install() {
    // SAFETY: registers a window class and creates a message-only window owned by this thread.
    let hwnd = unsafe {
        let Ok(hinst) = GetModuleHandleW(None) else {
            return;
        };
        let wc = WNDCLASSW {
            lpfnWndProc: Some(wake_proc),
            hInstance: hinst.into(),
            lpszClassName: WAKE_CLASS,
            ..Default::default()
        };
        RegisterClassW(&wc); // already registered is fine
        CreateWindowExW(
            WINDOW_EX_STYLE(0),
            WAKE_CLASS,
            w!(""),
            WINDOW_STYLE(0),
            0,
            0,
            0,
            0,
            Some(HWND_MESSAGE),
            None,
            Some(hinst.into()),
            None,
        )
    };
    let Ok(hwnd) = hwnd else {
        crate::log::note("cells: could not create the wake window; relying on the timer");
        return;
    };
    WAKE_WINDOW.with(|w| w.set(hwnd.0 as isize));
    let raw = hwnd.0 as isize;
    let wake = move || {
        // SAFETY: PostMessageW is thread-safe; a stale handle just fails.
        let _ = unsafe { PostMessageW(Some(HWND(raw as *mut _)), WM_APP, WPARAM(0), LPARAM(0)) };
    };
    xlapi::init(Bridge::new(
        Arc::new(execute),
        Box::new(wake),
        xlapi::CALL_TIMEOUT,
    ));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn values_survive_a_trip_through_variants() {
        for v in [
            Value::Empty,
            Value::Num(1.5),
            Value::Bool(true),
            Value::Str("héllo".into()),
            Value::Err(ErrorCode::NA),
            Value::Err(ErrorCode::Div0),
        ] {
            assert_eq!(value_of(&variant_of(&v)), v);
        }
    }

    #[test]
    fn tables_survive_a_trip_through_safearrays() {
        let a = Value::Array(Array::new(
            2,
            3,
            vec![
                Value::Num(1.0),
                Value::Str("b".into()),
                Value::Empty,
                Value::Bool(false),
                Value::Err(ErrorCode::Ref),
                Value::Num(6.0),
            ],
        ));
        let v = write_variant(&a).unwrap();
        assert_eq!(read_variant(&v).unwrap(), a);
    }

    #[test]
    fn error_codes_use_excels_scode_numbering() {
        assert_eq!(error_scode(ErrorCode::NA) as u32, 0x800A_07FA);
        assert_eq!(error_from_scode(0x800A_07D7u32 as i32), ErrorCode::Div0);
    }
}
