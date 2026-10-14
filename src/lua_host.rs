//! One Lua state per trusted workbook, plus Value <-> Lua conversion.
//!
//! A `LuaHost` only exists after trust has been granted; the blocked path never constructs one,
//! so untrusted scripts are never even compiled.
//!
//! Conventions for Lua code:
//! * numbers/strings/booleans map directly; an empty cell is `nil` (or `xl.EMPTY` inside arrays)
//! * errors are `xl.NA`, `xl.VALUE`, `xl.DIV0`, ... and may be returned or compared with `==`
//! * a range is a table of rows (`{{1,2},{3,4}}`); returning a flat list yields one column

use std::path::{Path, PathBuf};

use mlua::{Function, Lua, LuaOptions, MultiValue, StdLib, UserData, Value as LV};

use crate::package::ScriptSet;
use crate::value::{Array, ErrorCode, Value};

#[derive(Debug, thiserror::Error)]
pub enum HostError {
    #[error("{0}")]
    Lua(String),
    #[error("lua function '{0}' is not defined")]
    Undefined(String),
}

impl From<mlua::Error> for HostError {
    fn from(e: mlua::Error) -> Self {
        // Lua errors carry a traceback; the first line is what a spreadsheet cell can show.
        let msg = e.to_string();
        HostError::Lua(msg.lines().next().unwrap_or("lua error").to_owned())
    }
}

#[derive(Clone, Copy)]
struct XlError(ErrorCode);

impl UserData for XlError {
    fn add_methods<M: mlua::UserDataMethods<Self>>(methods: &mut M) {
        methods.add_meta_method("__tostring", |_, this, ()| {
            Ok(format!("#{}", this.0.lua_name()))
        });
        methods.add_meta_method("__eq", |_, this, other: mlua::AnyUserData| {
            Ok(other
                .borrow::<XlError>()
                .map(|o| o.0 == this.0)
                .unwrap_or(false))
        });
    }
}

struct XlEmpty;
impl UserData for XlEmpty {
    fn add_methods<M: mlua::UserDataMethods<Self>>(methods: &mut M) {
        methods.add_meta_method("__tostring", |_, _, ()| Ok(String::new()));
    }
}

pub struct HostContext<'a> {
    pub workbook: &'a Path,
}

pub struct LuaHost {
    lua: Lua,
}

impl LuaHost {
    /// Builds the state and runs every script once, defining its globals.
    ///
    /// The full standard library is available (io, os, package, ...) and C modules may be
    /// loaded: once a workbook is trusted its scripts are treated as the user's own code.
    pub fn new(scripts: &ScriptSet, ctx: &HostContext<'_>) -> Result<Self, HostError> {
        // SAFETY: enables loading C modules via `require`/`package.loadlib`. Reachable only
        // after the user explicitly trusted this exact script set (see trust.rs).
        let lua = unsafe { Lua::unsafe_new_with(StdLib::ALL_SAFE, LuaOptions::default()) };
        install_xl_table(&lua, ctx)?;
        configure_package_path(&lua, ctx.workbook)?;
        for s in &scripts.scripts {
            lua.load(&s.code).set_name(format!("={}", s.name)).exec()?;
        }
        Ok(LuaHost { lua })
    }

    pub fn call(&self, func: &str, args: &[Value]) -> Result<Value, HostError> {
        let f: LV = self.lua.globals().get(func)?;
        let LV::Function(f) = f else {
            return Err(HostError::Undefined(func.to_owned()));
        };
        self.invoke(&f, args)
    }

    /// Runs a chunk with the arguments available as `...`. A bare expression is accepted too.
    pub fn eval(&self, code: &str, args: &[Value]) -> Result<Value, HostError> {
        let f = match self
            .lua
            .load(format!("return {code}"))
            .set_name("=eval")
            .into_function()
        {
            Ok(f) => f,
            Err(_) => self.lua.load(code).set_name("=eval").into_function()?,
        };
        self.invoke(&f, args)
    }

    fn invoke(&self, f: &Function, args: &[Value]) -> Result<Value, HostError> {
        let lua_args: Vec<LV> = args
            .iter()
            .map(|a| to_lua(&self.lua, a))
            .collect::<Result<_, _>>()?;
        let ret: MultiValue = f.call(MultiValue::from_iter(lua_args))?;
        match ret.into_iter().next() {
            None => Ok(Value::Empty),
            Some(v) => Ok(from_lua(&v)?),
        }
    }
}

fn install_xl_table(lua: &Lua, ctx: &HostContext<'_>) -> mlua::Result<()> {
    let xl = lua.create_table()?;
    for e in ErrorCode::ALL {
        xl.set(e.lua_name(), lua.create_userdata(XlError(e))?)?;
    }
    xl.set("EMPTY", lua.create_userdata(XlEmpty)?)?;
    xl.set("version", env!("CARGO_PKG_VERSION"))?;
    xl.set("workbook", ctx.workbook.to_string_lossy().into_owned())?;
    lua.globals().set("xl", xl)
}

/// `require` finds modules next to the workbook, then in `%APPDATA%\xll-bridge\lua`.
fn configure_package_path(lua: &Lua, workbook: &Path) -> mlua::Result<()> {
    let mut dirs: Vec<PathBuf> = Vec::new();
    if let Some(d) = workbook.parent() {
        dirs.push(d.to_path_buf());
    }
    if let Some(p) =
        crate::trust::TrustStore::default_path().and_then(|p| p.parent().map(|d| d.join("lua")))
    {
        dirs.push(p);
    }
    let sep = std::path::MAIN_SEPARATOR;
    let lua_paths: Vec<String> = dirs
        .iter()
        .flat_map(|d| {
            let d = d.to_string_lossy();
            [format!("{d}{sep}?.lua"), format!("{d}{sep}?{sep}init.lua")]
        })
        .collect();
    let dll_ext = if cfg!(windows) { "dll" } else { "so" };
    let c_paths: Vec<String> = dirs
        .iter()
        .map(|d| format!("{}{sep}?.{dll_ext}", d.to_string_lossy()))
        .collect();

    let package: mlua::Table = lua.globals().get("package")?;
    let old_path: String = package.get("path")?;
    let old_cpath: String = package.get("cpath")?;
    package.set("path", format!("{};{old_path}", lua_paths.join(";")))?;
    package.set("cpath", format!("{};{old_cpath}", c_paths.join(";")))
}

pub fn to_lua(lua: &Lua, v: &Value) -> mlua::Result<LV> {
    Ok(match v {
        Value::Empty => LV::Nil,
        Value::Num(n) => LV::Number(*n),
        Value::Bool(b) => LV::Boolean(*b),
        Value::Str(s) => LV::String(lua.create_string(s)?),
        Value::Err(e) => {
            let xl: mlua::Table = lua.globals().get("xl")?;
            xl.get(e.lua_name())?
        }
        Value::Array(a) => {
            let xl: mlua::Table = lua.globals().get("xl")?;
            let empty: LV = xl.get("EMPTY")?;
            let rows = lua.create_table_with_capacity(a.rows, 0)?;
            for r in 0..a.rows {
                let row = lua.create_table_with_capacity(a.cols, 0)?;
                for c in 0..a.cols {
                    let cell = match &a.data[r * a.cols + c] {
                        Value::Empty => empty.clone(),
                        other => to_lua(lua, other)?,
                    };
                    row.raw_set(c + 1, cell)?;
                }
                rows.raw_set(r + 1, row)?;
            }
            LV::Table(rows)
        }
    })
}

pub fn from_lua(v: &LV) -> mlua::Result<Value> {
    Ok(match v {
        LV::Nil => Value::Empty,
        LV::Boolean(b) => Value::Bool(*b),
        LV::Integer(i) => Value::Num(*i as f64),
        LV::Number(n) => Value::Num(*n),
        LV::String(s) => Value::Str(s.to_string_lossy()),
        LV::UserData(u) => {
            if let Ok(e) = u.borrow::<XlError>() {
                Value::Err(e.0)
            } else if u.borrow::<XlEmpty>().is_ok() {
                Value::Empty
            } else {
                return Err(mlua::Error::runtime("cannot return this userdata to Excel"));
            }
        }
        LV::Table(t) => table_to_array(t)?,
        other => {
            return Err(mlua::Error::runtime(format!(
                "cannot return a {} to Excel",
                other.type_name()
            )));
        }
    })
}

fn table_to_array(t: &mlua::Table) -> mlua::Result<Value> {
    let len = t.raw_len();
    if len == 0 {
        return Ok(Value::Empty);
    }
    let first: LV = t.raw_get(1)?;
    let nested = matches!(first, LV::Table(_));
    let mut rows: Vec<Vec<Value>> = Vec::with_capacity(len);
    for i in 1..=len {
        let item: LV = t.raw_get(i)?;
        match (&item, nested) {
            (LV::Table(row), true) => {
                let n = row.raw_len();
                let mut cells = Vec::with_capacity(n);
                for j in 1..=n {
                    let c: LV = row.raw_get(j)?;
                    cells.push(from_lua(&c)?);
                }
                rows.push(cells);
            }
            (_, true) => {
                return Err(mlua::Error::runtime(
                    "mixed rows and scalars in returned table",
                ));
            }
            (_, false) => rows.push(vec![from_lua(&item)?]),
        }
    }
    let cols = rows.iter().map(Vec::len).max().unwrap_or(0).max(1);
    let mut data = Vec::with_capacity(rows.len() * cols);
    for r in &rows {
        data.extend(r.iter().cloned());
        data.extend(std::iter::repeat_n(Value::Empty, cols - r.len()));
    }
    Ok(Value::Array(Array::new(rows.len(), cols, data)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::Script;

    fn host(code: &str) -> LuaHost {
        let set = ScriptSet {
            scripts: vec![Script {
                name: "t".into(),
                code: code.into(),
            }],
        };
        LuaHost::new(
            &set,
            &HostContext {
                workbook: Path::new("/tmp/book.xlsx"),
            },
        )
        .unwrap()
    }

    #[test]
    fn scalar_call() {
        let h = host("function add(a, b) return a + b end");
        assert_eq!(
            h.call("add", &[Value::Num(2.0), Value::Num(3.0)]).unwrap(),
            Value::Num(5.0)
        );
    }

    #[test]
    fn undefined_function() {
        let h = host("");
        assert!(matches!(h.call("nope", &[]), Err(HostError::Undefined(_))));
    }

    #[test]
    fn ranges_in_and_out() {
        let h = host(
            "function transpose(r) local o = {} for i = 1, #r[1] do o[i] = {} for j = 1, #r do o[i][j] = r[j][i] end end return o end",
        );
        let input = Value::Array(Array::new(
            2,
            3,
            (1..=6).map(|n| Value::Num(n as f64)).collect(),
        ));
        let out = h.call("transpose", &[input]).unwrap();
        let Value::Array(a) = out else {
            panic!("expected array")
        };
        assert_eq!((a.rows, a.cols), (3, 2));
        assert_eq!(a.data[1], Value::Num(4.0));
    }

    #[test]
    fn flat_list_is_a_column_and_ragged_rows_pad() {
        let h = host("function col() return {1, 2, 3} end function rag() return {{1}, {2, 3}} end");
        let Value::Array(c) = h.call("col", &[]).unwrap() else {
            panic!()
        };
        assert_eq!((c.rows, c.cols), (3, 1));
        let Value::Array(r) = h.call("rag", &[]).unwrap() else {
            panic!()
        };
        assert_eq!((r.rows, r.cols), (2, 2));
        assert_eq!(r.data[1], Value::Empty);
    }

    #[test]
    fn errors_and_empties_round_trip() {
        let h = host(
            "function id(x) return x end function na() return xl.NA end function chk(x) return x == xl.DIV0 end",
        );
        assert_eq!(
            h.call("id", &[Value::Err(ErrorCode::Ref)]).unwrap(),
            Value::Err(ErrorCode::Ref)
        );
        assert_eq!(h.call("na", &[]).unwrap(), Value::Err(ErrorCode::NA));
        assert_eq!(
            h.call("chk", &[Value::Err(ErrorCode::Div0)]).unwrap(),
            Value::Bool(true)
        );
        assert_eq!(h.call("id", &[Value::Empty]).unwrap(), Value::Empty);
    }

    #[test]
    fn eval_accepts_expressions_and_statements() {
        let h = host("");
        assert_eq!(
            h.eval("1 + ...", &[Value::Num(1.0)]).unwrap(),
            Value::Num(2.0)
        );
        assert_eq!(
            h.eval("local s = 0 for i=1,4 do s = s + i end return s", &[])
                .unwrap(),
            Value::Num(10.0)
        );
    }

    #[test]
    fn lua_errors_surface_as_one_line() {
        let h = host("function boom() error('bad thing') end");
        let Err(HostError::Lua(m)) = h.call("boom", &[]) else {
            panic!()
        };
        assert!(m.contains("bad thing") && !m.contains('\n'));
    }

    #[test]
    fn can_run_external_programs() {
        if cfg!(windows) {
            return;
        }
        let h = host(
            "function sh() local p = io.popen('echo hi') local s = p:read('*l') p:close() return s end",
        );
        assert_eq!(h.call("sh", &[]).unwrap(), Value::Str("hi".into()));
    }
}
