//! Raw Excel C API (XLOPER12) definitions and the `MdCallBack12` trampoline.
//!
//! Excel exports `MdCallBack12` from its own process, so the XLL needs neither the Microsoft
//! XLL SDK headers nor `xlcall32.lib`: we resolve the entry point at load time.
//! Layout and constants match XLCALL.H (x64 only; the 32-bit `__stdcall` name decoration
//! of exports is not handled).

#![allow(non_upper_case_globals, dead_code)]

use std::sync::atomic::{AtomicUsize, Ordering};

use windows_sys::Win32::System::LibraryLoader::{GetModuleHandleA, GetProcAddress};

pub type RW = i32;
#[allow(clippy::upper_case_acronyms)]
pub type COL = i32;

#[repr(C)]
#[derive(Clone, Copy, Default)]
pub struct XlRef12 {
    pub rw_first: RW,
    pub rw_last: RW,
    pub col_first: COL,
    pub col_last: COL,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct SRef {
    pub count: u16,
    pub reference: XlRef12,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct MultiArray {
    pub lparray: *mut Xloper12,
    pub rows: RW,
    pub columns: COL,
}

/// `val` union. `_pad` pins the size/alignment to the SDK's 24 bytes (its `flow` and
/// `bigdata` members are never read here).
#[repr(C)]
#[derive(Clone, Copy)]
pub union XlVal {
    pub num: f64,
    pub str_: *mut u16,
    pub xbool: i32,
    pub err: i32,
    pub w: i32,
    pub sref: SRef,
    pub array: MultiArray,
    _pad: [u64; 3],
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct Xloper12 {
    pub val: XlVal,
    pub xltype: u32,
}

const _: () = {
    assert!(std::mem::size_of::<Xloper12>() == 32);
    assert!(std::mem::offset_of!(Xloper12, xltype) == 24);
};

pub const xltypeNum: u32 = 0x0001;
pub const xltypeStr: u32 = 0x0002;
pub const xltypeBool: u32 = 0x0004;
pub const xltypeRef: u32 = 0x0008;
pub const xltypeErr: u32 = 0x0010;
pub const xltypeMulti: u32 = 0x0040;
pub const xltypeMissing: u32 = 0x0080;
pub const xltypeNil: u32 = 0x0100;
pub const xltypeSRef: u32 = 0x0400;
pub const xltypeInt: u32 = 0x0800;
pub const xlbitXLFree: u32 = 0x1000;
pub const xlbitDLLFree: u32 = 0x4000;
pub const XLTYPE_MASK: u32 = 0x0FFF;

pub const xlretSuccess: i32 = 0;

const xlCommand: i32 = 0x8000;
const xlSpecial: i32 = 0x4000;
pub const xlFree: i32 = xlSpecial;
pub const xlCoerce: i32 = 2 | xlSpecial;
pub const xlGetHwnd: i32 = 8 | xlSpecial;
pub const xlSheetNm: i32 = 5 | xlSpecial;
pub const xlGetName: i32 = 9 | xlSpecial;
pub const xlfCaller: i32 = 89;
pub const xlfCell: i32 = 125;
pub const xlfRegister: i32 = 149;
pub const xlfUnregister: i32 = 201;
pub const xlcCalculateNow: i32 = 141 | xlCommand;
pub const xlcCalculateDocument: i32 = 167 | xlCommand;

pub const xlerrNA: i32 = 42;

type MdCallBack12 = unsafe extern "system" fn(
    xlfn: i32,
    coper: i32,
    args: *const *mut Xloper12,
    res: *mut Xloper12,
) -> i32;

static ENTRY: AtomicUsize = AtomicUsize::new(0);

fn entry() -> Option<MdCallBack12> {
    let mut p = ENTRY.load(Ordering::Relaxed);
    if p == 0 {
        // SAFETY: plain Win32 lookups; the name is NUL-terminated.
        let found = unsafe {
            let exe = GetModuleHandleA(std::ptr::null());
            GetProcAddress(exe, c"MdCallBack12".as_ptr().cast())
        };
        p = found.map_or(0, |f| f as usize);
        ENTRY.store(p, Ordering::Relaxed);
    }
    // SAFETY: the address came from Excel's export of exactly this signature.
    (p != 0).then(|| unsafe { std::mem::transmute::<usize, MdCallBack12>(p) })
}

impl Xloper12 {
    pub fn nil() -> Self {
        Xloper12 {
            val: XlVal { _pad: [0; 3] },
            xltype: xltypeNil,
        }
    }

    pub fn missing() -> Self {
        Xloper12 {
            val: XlVal { _pad: [0; 3] },
            xltype: xltypeMissing,
        }
    }

    pub fn int(w: i32) -> Self {
        let mut v = XlVal { _pad: [0; 3] };
        v.w = w;
        Xloper12 {
            val: v,
            xltype: xltypeInt,
        }
    }

    pub fn base_type(&self) -> u32 {
        self.xltype & XLTYPE_MASK
    }

    /// Borrowed string operand. `buf` must be a counted UTF-16 string (length in `buf[0]`)
    /// that outlives every use of the returned value.
    pub fn borrowed_str(buf: &[u16]) -> Self {
        let mut v = XlVal { _pad: [0; 3] };
        v.str_ = buf.as_ptr() as *mut u16;
        Xloper12 {
            val: v,
            xltype: xltypeStr,
        }
    }
}

/// Counted UTF-16 string as Excel wants it: `[len, chars...]`, at most 32767 code units.
pub fn counted_utf16(s: &str) -> Vec<u16> {
    let mut units: Vec<u16> = s.encode_utf16().take(32767).collect();
    units.insert(0, units.len() as u16);
    units
}

/// Reads an Excel-owned counted string.
///
/// # Safety
/// `p` must point at a valid counted UTF-16 string.
pub unsafe fn read_counted(p: *const u16) -> String {
    if p.is_null() {
        return String::new();
    }
    let len = unsafe { *p } as usize;
    String::from_utf16_lossy(unsafe { std::slice::from_raw_parts(p.add(1), len) })
}

/// Calls the Excel C API. Returns the `xlret*` code; `res` is filled on success and, when it
/// carries `xlbitXLFree`, must be passed to [`free`].
pub fn excel12(xlfn: i32, res: Option<&mut Xloper12>, args: &mut [&mut Xloper12]) -> i32 {
    let Some(f) = entry() else { return -1 };
    let ptrs: Vec<*mut Xloper12> = args.iter_mut().map(|a| &mut **a as *mut Xloper12).collect();
    let res_ptr = res.map_or(std::ptr::null_mut(), |r| r as *mut Xloper12);
    // SAFETY: pointers are valid for the duration of the call; Excel's contract.
    unsafe { f(xlfn, ptrs.len() as i32, ptrs.as_ptr(), res_ptr) }
}

/// Releases memory Excel allocated for a result (`xlbitXLFree`).
pub fn free(x: &mut Xloper12) {
    if x.xltype & xlbitXLFree != 0 {
        let _ = excel12(xlFree, None, &mut [x]);
    }
}
