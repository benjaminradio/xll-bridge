//! Excel-side value model, independent of the XLOPER12 ABI so it can be tested anywhere.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ErrorCode {
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NA,
}

impl ErrorCode {
    pub const ALL: [ErrorCode; 7] = [
        ErrorCode::Null,
        ErrorCode::Div0,
        ErrorCode::Value,
        ErrorCode::Ref,
        ErrorCode::Name,
        ErrorCode::Num,
        ErrorCode::NA,
    ];

    /// The `xlerr*` constant from XLCALL.H.
    pub fn xl_code(self) -> i32 {
        match self {
            ErrorCode::Null => 0,
            ErrorCode::Div0 => 7,
            ErrorCode::Value => 15,
            ErrorCode::Ref => 23,
            ErrorCode::Name => 29,
            ErrorCode::Num => 36,
            ErrorCode::NA => 42,
        }
    }

    pub fn from_xl_code(code: i32) -> Option<ErrorCode> {
        Self::ALL.into_iter().find(|e| e.xl_code() == code)
    }

    /// Name as exposed to Lua (`xl.NA`, `xl.DIV0`, ...).
    pub fn lua_name(self) -> &'static str {
        match self {
            ErrorCode::Null => "NULL",
            ErrorCode::Div0 => "DIV0",
            ErrorCode::Value => "VALUE",
            ErrorCode::Ref => "REF",
            ErrorCode::Name => "NAME",
            ErrorCode::Num => "NUM",
            ErrorCode::NA => "NA",
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    Empty,
    Num(f64),
    Bool(bool),
    Str(String),
    Err(ErrorCode),
    /// Row-major 2D block.
    Array(Array),
}

#[derive(Clone, Debug, PartialEq)]
pub struct Array {
    pub rows: usize,
    pub cols: usize,
    pub data: Vec<Value>,
}

impl Array {
    pub fn new(rows: usize, cols: usize, data: Vec<Value>) -> Self {
        assert_eq!(rows * cols, data.len(), "array shape does not match data");
        Array { rows, cols, data }
    }
}
