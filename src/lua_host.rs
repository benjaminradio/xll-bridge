//! One Lua state per trusted workbook, plus Value <-> Lua conversion.
//!
//! A `LuaHost` only exists after trust has been granted; the blocked path never constructs one,
//! so untrusted scripts are never even compiled.
//!
//! Conventions for Lua code:
//! * numbers/strings/booleans map directly; an empty cell is `nil` (or `xl.EMPTY` inside arrays)
//! * errors are `xl.NA`, `xl.VALUE`, `xl.DIV0`, ... and may be returned or compared with `==`
//! * a range is a table of rows (`{{1,2},{3,4}}`); returning a flat list yields one column

use std::cell::{Cell, RefCell};
use std::path::{Path, PathBuf};
use std::rc::Rc;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::{Duration, Instant};

use mlua::chunk::ChunkMode;
use mlua::{
    Function, HookTriggers, Lua, LuaOptions, MultiValue, StdLib, UserData, Value as LV, VmState,
};

use crate::procrun::{self, RunSpec};
use crate::sqldb;
use crate::value::{Array, ErrorCode, Value};
use crate::vfs::Bundle;
use crate::xlapi::{self, MAX_CELLS, Op, Reply, XlApi};

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
    /// Where native modules (`.dll`/`.so`) from the embedded tree are extracted so `require` can
    /// load them. `None` disables extraction.
    pub cache_dir: Option<&'a Path>,
    /// Access to worksheet cells for `xl.range` and friends; `None` when Excel is not attached
    /// (tests, `xllb-pack`), in which case those functions raise an error.
    pub xl: Option<Arc<dyn XlApi>>,
}

/// VM instructions a sandboxed call may run before it is aborted (a few seconds of work).
const SANDBOX_INSTRUCTION_LIMIT: u64 = 200_000_000;
const SANDBOX_HOOK_STEP: u32 = 100_000;
const SANDBOX_MEMORY_LIMIT: usize = 64 * 1024 * 1024;

/// Bounds for one background call: a deadline, and a flag that aborts it from outside.
#[derive(Clone)]
pub struct Limits {
    pub deadline: Instant,
    pub cancel: Arc<AtomicBool>,
}

/// Instructions between checks of [`Limits`] in a worker state.
const LIMIT_CHECK_STEP: u32 = 10_000;

pub struct LuaHost {
    lua: Lua,
    /// Present for sandboxed states: instructions executed in the current call.
    budget: Option<Rc<Cell<u64>>>,
    /// Consulted by the hook of worker states; empty outside [`LuaHost::call_limited`].
    limits: Rc<RefCell<Option<Limits>>>,
}

impl LuaHost {
    /// Builds the state and runs the root-level `.lua` files once, in name order, defining
    /// their globals. Files in folders are available to `require` and `xl.read_file`.
    ///
    /// The full standard library is available (io, os, package, ...) and C modules may be
    /// loaded: once a workbook is trusted its files are treated as the user's own code.
    pub fn new(bundle: &Bundle, ctx: &HostContext<'_>) -> Result<Self, HostError> {
        Self::new_with(bundle, ctx, false)
    }

    /// Like [`LuaHost::new`] for a background worker: the state also honours the [`Limits`] given
    /// to [`LuaHost::call_limited`], so a runaway or cancelled call is aborted.
    pub fn new_worker(bundle: &Bundle, ctx: &HostContext<'_>) -> Result<Self, HostError> {
        Self::new_with(bundle, ctx, true)
    }

    fn new_with(bundle: &Bundle, ctx: &HostContext<'_>, worker: bool) -> Result<Self, HostError> {
        // SAFETY: enables loading C modules via `require`/`package.loadlib`. Reachable only
        // after the user explicitly trusted this exact file set (see trust.rs).
        let lua = unsafe { Lua::unsafe_new_with(StdLib::ALL_SAFE, LuaOptions::default()) };
        install_xl_table(&lua, Some(ctx.workbook))?;
        let shared = Rc::new(bundle.clone());
        install_embedded_fs(&lua, &shared)?;
        let limits: Rc<RefCell<Option<Limits>>> = Rc::new(RefCell::new(None));
        install_process_api(&lua, Rc::clone(&limits))?;
        install_range_api(&lua, ctx.workbook, ctx.xl.clone())?;
        install_macro_api(&lua)?;
        install_region_api(&lua, ctx.workbook, ctx.xl.clone())?;
        install_file_helpers(&lua)?;
        let sql_env = Rc::new(SqlEnv {
            api: ctx.xl.clone(),
            workbook: ctx.workbook.to_owned(),
            bundle: Arc::new(bundle.clone()),
            limits: Rc::clone(&limits),
        });
        install_sql_api(&lua, Rc::clone(&sql_env))?;
        install_db_api(&lua, sql_env)?;
        if worker {
            let hook_limits = Rc::clone(&limits);
            lua.set_hook(
                HookTriggers::new().every_nth_instruction(LIMIT_CHECK_STEP),
                move |_, _| {
                    if let Some(l) = hook_limits.borrow().as_ref() {
                        if l.cancel.load(Ordering::Relaxed) {
                            return Err(mlua::Error::runtime("cancelled"));
                        }
                        if Instant::now() >= l.deadline {
                            return Err(mlua::Error::runtime("timed out"));
                        }
                    }
                    Ok(VmState::Continue)
                },
            )?;
        }
        if worker {
            install_sleep(&lua, Rc::clone(&limits))?;
        }
        let native = match ctx.cache_dir {
            Some(dir) => extract_native(bundle, dir)
                .map_err(|e| HostError::Lua(format!("cannot extract native modules: {e}")))?,
            None => None,
        };
        configure_package_path(&lua, ctx.workbook, native.as_deref())?;
        // Machine scripts (the root `.lua` files of %APPDATA%\xll-bridge\local) come first, so a
        // workbook's own scripts can use what they define.
        for f in crate::dbstore::machine_scripts() {
            lua.load(strip_bom(&f.data))
                .set_name(format!("=local/{}", f.name))
                .set_mode(ChunkMode::Text)
                .exec()?;
        }
        for f in bundle.root_scripts() {
            lua.load(strip_bom(&f.data))
                .set_name(format!("={}", f.path))
                .set_mode(ChunkMode::Text)
                .exec()?;
        }
        Ok(LuaHost {
            lua,
            budget: None,
            limits,
        })
    }

    /// A state for expressions typed into cells (`LUA.EVAL`) in workbooks that carry no trusted
    /// scripts. Pure computation only: no io/os/package/debug, no `load`/`dofile`/`loadfile`,
    /// bounded memory and instruction count. Safe to run without any trust decision.
    pub fn sandboxed() -> Result<Self, HostError> {
        let libs = StdLib::COROUTINE | StdLib::TABLE | StdLib::STRING | StdLib::UTF8 | StdLib::MATH;
        let lua = Lua::new_with(libs, LuaOptions::default())?;
        let globals = lua.globals();
        for name in ["dofile", "loadfile", "load", "loadstring", "require"] {
            globals.set(name, LV::Nil)?;
        }
        lua.set_memory_limit(SANDBOX_MEMORY_LIMIT)?;
        install_xl_table(&lua, None)?;
        let budget = Rc::new(Cell::new(0u64));
        let hook_budget = Rc::clone(&budget);
        lua.set_hook(
            HookTriggers::new().every_nth_instruction(SANDBOX_HOOK_STEP),
            move |_, _| {
                let used = hook_budget.get() + u64::from(SANDBOX_HOOK_STEP);
                hook_budget.set(used);
                if used > SANDBOX_INSTRUCTION_LIMIT {
                    Err(mlua::Error::runtime(
                        "expression exceeded the instruction limit",
                    ))
                } else {
                    Ok(VmState::Continue)
                }
            },
        )?;
        Ok(LuaHost {
            lua,
            budget: Some(budget),
            limits: Rc::new(RefCell::new(None)),
        })
    }

    pub fn call(&self, func: &str, args: &[Value]) -> Result<Value, HostError> {
        let f: LV = self.lua.globals().get(func)?;
        let LV::Function(f) = f else {
            return Err(HostError::Undefined(func.to_owned()));
        };
        self.invoke(&f, args)
    }

    /// [`LuaHost::call`] under `limits` (a worker state; see [`LuaHost::new_worker`]).
    pub fn call_limited(
        &self,
        func: &str,
        args: &[Value],
        limits: &Limits,
    ) -> Result<Value, HostError> {
        *self.limits.borrow_mut() = Some(limits.clone());
        let result = self.call(func, args);
        *self.limits.borrow_mut() = None;
        result
    }

    /// Names registered with `xl.macro(name, fn)`, sorted.
    pub fn macro_names(&self) -> Vec<String> {
        let Ok(xl) = self.lua.globals().get::<mlua::Table>("xl") else {
            return Vec::new();
        };
        let Ok(macros) = xl.get::<mlua::Table>("macros") else {
            return Vec::new();
        };
        let mut names: Vec<String> = macros
            .pairs::<String, LV>()
            .filter_map(|p| p.ok().map(|(k, _)| k))
            .collect();
        names.sort();
        names
    }

    /// Runs the macro `name` (a worker state; see [`LuaHost::new_worker`]) under `limits`.
    pub fn call_macro(&self, name: &str, limits: &Limits) -> Result<(), HostError> {
        let xl: mlua::Table = self.lua.globals().get("xl")?;
        let macros: mlua::Table = xl.get("macros")?;
        let LV::Function(f) = macros.get::<LV>(name)? else {
            return Err(HostError::Undefined(name.to_owned()));
        };
        *self.limits.borrow_mut() = Some(limits.clone());
        let result = f.call::<MultiValue>(()).map(drop);
        *self.limits.borrow_mut() = None;
        Ok(result?)
    }

    /// Runs a streaming function (a worker state): the function pushes values with `xl.emit(v)`
    /// for as long as it runs, and a value it returns at the end is pushed too. Only scalars
    /// can be pushed.
    pub fn run_stream(
        &self,
        func: &str,
        args: &[Value],
        sink: Arc<dyn Fn(Value) + Send + Sync>,
        limits: &Limits,
    ) -> Result<(), HostError> {
        let xl: mlua::Table = self.lua.globals().get("xl")?;
        let emit_sink = Arc::clone(&sink);
        let emit = self.lua.create_function(move |_, v: LV| {
            let value = from_lua(&v)?;
            if matches!(value, Value::Array(_)) {
                return Err(mlua::Error::runtime(
                    "a stream can only emit single values, not tables",
                ));
            }
            emit_sink(value);
            Ok(())
        })?;
        xl.set("emit", emit)?;
        *self.limits.borrow_mut() = Some(limits.clone());
        let result = self.call(func, args);
        *self.limits.borrow_mut() = None;
        xl.set("emit", LV::Nil)?;
        match result? {
            Value::Empty => {}
            last => sink(last),
        }
        Ok(())
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
        if let Some(b) = &self.budget {
            b.set(0);
        }
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

fn install_xl_table(lua: &Lua, workbook: Option<&Path>) -> mlua::Result<()> {
    let xl = lua.create_table()?;
    for e in ErrorCode::ALL {
        xl.set(e.lua_name(), lua.create_userdata(XlError(e))?)?;
    }
    xl.set("EMPTY", lua.create_userdata(XlEmpty)?)?;
    xl.set("version", env!("CARGO_PKG_VERSION"))?;
    // Excel date numbers <-> ISO 8601 text (nil when it cannot be converted).
    xl.set(
        "date_to_iso",
        lua.create_function(|_, n: f64| Ok(crate::dates::serial_to_iso(n)))?,
    )?;
    xl.set(
        "iso_to_date",
        lua.create_function(|_, s: String| Ok(crate::dates::iso_to_serial(&s)))?,
    )?;
    if let Some(wb) = workbook {
        xl.set("workbook", wb.to_string_lossy().into_owned())?;
    }
    lua.globals().set("xl", xl)
}

/// A worksheet range as seen from Lua: `xl.range("Data!A1:C10")`. The address is resolved once,
/// when the object is made, so later calls cannot silently move to another sheet.
struct LuaRange {
    api: Arc<dyn XlApi>,
    workbook: PathBuf,
    address: String,
    rows: u32,
    cols: u32,
}

fn api_err(e: xlapi::XlError) -> mlua::Error {
    mlua::Error::runtime(e.message)
}

impl LuaRange {
    fn read(&self, lua: &Lua, formulas: bool, always_table: bool) -> mlua::Result<LV> {
        let reply = self
            .api
            .call(Op::Read {
                workbook: self.workbook.clone(),
                address: self.address.clone(),
                formulas,
            })
            .map_err(api_err)?;
        let Reply::Data(v) = reply else {
            return Err(mlua::Error::runtime("unexpected reply from Excel"));
        };
        if !always_table
            && let Value::Array(a) = &v
            && a.data.len() == 1
        {
            return to_lua(lua, &a.data[0]);
        }
        to_lua(lua, &v)
    }

    fn write(&self, v: &LV, formulas: bool) -> mlua::Result<()> {
        let data = from_lua(v)?;
        let cells = match &data {
            Value::Array(a) => (a.rows as u64) * (a.cols as u64),
            _ => 1,
        };
        if cells > MAX_CELLS {
            return Err(mlua::Error::runtime(format!(
                "too many cells to write at once ({cells}; the limit is {MAX_CELLS})"
            )));
        }
        self.api
            .call(Op::Write {
                workbook: self.workbook.clone(),
                address: self.address.clone(),
                formulas,
                data,
            })
            .map_err(api_err)?;
        Ok(())
    }
}

impl UserData for LuaRange {
    fn add_methods<M: mlua::UserDataMethods<Self>>(methods: &mut M) {
        // A single cell gives a bare value, anything larger a table of rows.
        methods.add_method("value", |lua, this, ()| this.read(lua, false, false));
        // Always a table of rows, even for one cell.
        methods.add_method("values", |lua, this, ()| this.read(lua, false, true));
        methods.add_method("formulas", |lua, this, ()| this.read(lua, true, true));
        methods.add_method("set", |_, this, v: LV| this.write(&v, false));
        methods.add_method("set_formulas", |_, this, v: LV| this.write(&v, true));
        methods.add_method("clear", |_, this, ()| {
            this.api
                .call(Op::Clear {
                    workbook: this.workbook.clone(),
                    address: this.address.clone(),
                })
                .map_err(api_err)?;
            Ok(())
        });
        // range:as_table(name): make the data at (or around) this range an Excel table, so a
        // PivotTable can be built on it. A single cell means the block of data around it.
        methods.add_method("as_table", |_, this, name: String| {
            this.api
                .call(Op::AsTable {
                    workbook: this.workbook.clone(),
                    address: this.address.clone(),
                    name,
                })
                .map_err(api_err)?;
            Ok(())
        });
        methods.add_method("size", |_, this, ()| Ok((this.rows, this.cols)));
        methods.add_method("address", |_, this, ()| Ok(this.address.clone()));
        methods.add_meta_method("__tostring", |_, this, ()| {
            Ok(format!("range({})", this.address))
        });
    }
}

fn make_range(api: &Arc<dyn XlApi>, workbook: &Path, op: Op) -> mlua::Result<LuaRange> {
    let Reply::Info(info) = api.call(op).map_err(api_err)? else {
        return Err(mlua::Error::runtime("unexpected reply from Excel"));
    };
    if u64::from(info.rows) * u64::from(info.cols) > MAX_CELLS {
        return Err(mlua::Error::runtime(format!(
            "{} is too large ({} x {} cells; the limit is {MAX_CELLS} cells)",
            info.address, info.rows, info.cols
        )));
    }
    Ok(LuaRange {
        api: Arc::clone(api),
        workbook: workbook.to_owned(),
        address: info.address,
        rows: info.rows,
        cols: info.cols,
    })
}

/// `xl.range`, `xl.sheets`, `xl.used_range`, `xl.selection`.
fn install_range_api(lua: &Lua, workbook: &Path, api: Option<Arc<dyn XlApi>>) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;
    let api = api.unwrap_or_else(|| Arc::new(Detached));
    let wb = workbook.to_owned();

    let (a, w) = (Arc::clone(&api), wb.clone());
    xl.set(
        "range",
        lua.create_function(move |_, address: String| {
            make_range(
                &a,
                &w,
                Op::Info {
                    workbook: w.clone(),
                    address,
                },
            )
        })?,
    )?;

    let (a, w) = (Arc::clone(&api), wb.clone());
    xl.set(
        "sheets",
        lua.create_function(move |lua, ()| {
            let Reply::Names(names) = a
                .call(Op::Sheets {
                    workbook: w.clone(),
                })
                .map_err(api_err)?
            else {
                return Err(mlua::Error::runtime("unexpected reply from Excel"));
            };
            lua.create_sequence_from(names)
        })?,
    )?;

    let (a, w) = (Arc::clone(&api), wb.clone());
    xl.set(
        "used_range",
        lua.create_function(move |_, sheet: String| {
            make_range(
                &a,
                &w,
                Op::UsedRange {
                    workbook: w.clone(),
                    sheet,
                },
            )
        })?,
    )?;

    let (a, w) = (api, wb);
    xl.set(
        "selection",
        lua.create_function(move |_, ()| {
            make_range(
                &a,
                &w,
                Op::Selection {
                    workbook: w.clone(),
                },
            )
        })?,
    )
}

/// `xl.sleep(seconds)` for background states: sleeps in short slices so that cancelling or the
/// deadline end it promptly.
fn install_sleep(lua: &Lua, limits: Rc<RefCell<Option<Limits>>>) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;
    xl.set(
        "sleep",
        lua.create_function(move |_, seconds: f64| {
            let end = Instant::now() + Duration::from_secs_f64(seconds.clamp(0.0, 86_400.0));
            loop {
                if let Some(l) = limits.borrow().as_ref() {
                    if l.cancel.load(Ordering::Relaxed) {
                        return Err(mlua::Error::runtime("cancelled"));
                    }
                    if Instant::now() >= l.deadline {
                        return Err(mlua::Error::runtime("timed out"));
                    }
                }
                let now = Instant::now();
                if now >= end {
                    return Ok(());
                }
                std::thread::sleep((end - now).min(Duration::from_millis(50)));
            }
        })?,
    )
}

/// `xl.macro(name, fn)`: registers a function that the ribbon's Macros menu can run. Macros run
/// in the background like `LUA.ASYNC` calls but may change cells.
fn install_macro_api(lua: &Lua) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;
    let macros = lua.create_table()?;
    xl.set("macros", macros.clone())?;
    xl.set(
        "macro",
        lua.create_function(move |_, (name, f): (String, Function)| {
            if name.trim().is_empty() {
                return Err(mlua::Error::runtime("a macro needs a name"));
            }
            macros.set(name, f)
        })?,
    )
}

/// What SQL run from Lua needs: where bracketed references read from, the databases `[db:name]`
/// and `xl.db` can find, and how a running call can stop a query.
struct SqlEnv {
    api: Option<Arc<dyn XlApi>>,
    workbook: PathBuf,
    bundle: Arc<Bundle>,
    limits: Rc<RefCell<Option<Limits>>>,
}

impl SqlEnv {
    fn source(&self) -> Box<dyn sqldb::Source> {
        let base: Box<dyn sqldb::Source> = match &self.api {
            Some(a) => Box::new(sqldb::XlSource {
                api: Arc::clone(a),
                workbook: self.workbook.clone(),
            }),
            None => Box::new(sqldb::NoSource),
        };
        Box::new(sqldb::WithDatabases {
            base,
            databases: crate::dbstore::resolver(
                self.workbook.clone(),
                Some(Arc::clone(&self.bundle)),
            ),
        })
    }

    fn limits(&self) -> sqldb::Limits {
        sqldb::Limits {
            cancel: self.limits.borrow().as_ref().map(|l| Arc::clone(&l.cancel)),
            ..Default::default()
        }
    }
}

type SqlParams = (Vec<Value>, Vec<(String, Value)>);

/// A Lua table of parameters: array part -> `?1..`, string keys -> `:name`.
fn split_params(t: Option<mlua::Table>) -> mlua::Result<SqlParams> {
    let mut positional: Vec<Value> = Vec::new();
    let mut named = Vec::new();
    if let Some(t) = t {
        for pair in t.pairs::<LV, LV>() {
            let (k, v) = pair?;
            match k {
                LV::Integer(i) if i >= 1 => {
                    let i = i as usize;
                    if positional.len() < i {
                        positional.resize(i, Value::Empty);
                    }
                    positional[i - 1] = from_lua(&v)?;
                }
                LV::String(s) => named.push((s.to_string_lossy(), from_lua(&v)?)),
                _ => {}
            }
        }
    }
    Ok((positional, named))
}

fn records_to_lua(lua: &Lua, out: &sqldb::Output) -> mlua::Result<mlua::Table> {
    let list = lua.create_table_with_capacity(out.rows.len(), 0)?;
    for (i, row) in out.rows.iter().enumerate() {
        let rec = lua.create_table_with_capacity(0, out.columns.len())?;
        for (name, v) in out.columns.iter().zip(row) {
            rec.set(name.as_str(), to_lua(lua, v)?)?;
        }
        list.raw_set(i + 1, rec)?;
    }
    Ok(list)
}

/// `xl.sql`, `xl.sql_records`, `xl.sql_exec` and the positional `xl.query` (what `SQL.ASYNC` runs).
fn install_sql_api(lua: &Lua, env: Rc<SqlEnv>) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;
    let fail = mlua::Error::runtime;

    let e = Rc::clone(&env);
    xl.set(
        "sql",
        lua.create_function(move |lua, (sql, params): (String, Option<mlua::Table>)| {
            let (args, named) = split_params(params)?;
            let source = e.source();
            let out = sqldb::run(&sqldb::Query {
                sql: &sql,
                args: &args,
                named: &named,
                source: source.as_ref(),
                read_only: false,
                limits: e.limits(),
            })
            .map_err(fail)?;
            to_lua(lua, &out.with_header())
        })?,
    )?;

    let e = Rc::clone(&env);
    xl.set(
        "sql_records",
        lua.create_function(move |lua, (sql, params): (String, Option<mlua::Table>)| {
            let (args, named) = split_params(params)?;
            let source = e.source();
            let out = sqldb::run(&sqldb::Query {
                sql: &sql,
                args: &args,
                named: &named,
                source: source.as_ref(),
                read_only: false,
                limits: e.limits(),
            })
            .map_err(fail)?;
            records_to_lua(lua, &out)
        })?,
    )?;

    let e = Rc::clone(&env);
    xl.set(
        "sql_exec",
        lua.create_function(move |_, script: String| {
            let source = e.source();
            sqldb::exec_script(&script, source.as_ref(), &e.limits()).map_err(fail)
        })?,
    )?;

    // xl.query(sql, a1, a2, ...): worksheet-function conventions (tables t1.., parameters ?1..),
    // statements may only read. SQL.ASYNC calls it as the global `__sql`.
    let e = env;
    let query = lua.create_function(move |lua, (sql, rest): (String, MultiValue)| {
        let args: Vec<Value> = rest
            .iter()
            .map(from_lua)
            .collect::<mlua::Result<Vec<_>>>()?;
        let source = e.source();
        let out = sqldb::run(&sqldb::Query {
            sql: &sql,
            args: &args,
            named: &[],
            source: source.as_ref(),
            read_only: true,
            limits: e.limits(),
        })
        .map_err(fail)?;
        to_lua(lua, &out.with_header())
    })?;
    lua.globals().set("__sql", query.clone())?;
    xl.set("query", query)
}

fn string_list(t: &mlua::Table, key: &str) -> mlua::Result<Vec<String>> {
    Ok(match t.get::<LV>(key)? {
        LV::Table(l) => l
            .sequence_values::<String>()
            .collect::<mlua::Result<Vec<_>>>()?,
        LV::String(s) => vec![s.to_string_lossy()],
        _ => Vec::new(),
    })
}

/// `{ clean_names = true, dates = {...}, text = {...}, index = {...} }`
fn csv_options(opts: Option<mlua::Table>) -> mlua::Result<crate::csvio::Options> {
    let mut o = crate::csvio::Options {
        clean_names: true,
        ..Default::default()
    };
    if let Some(t) = opts {
        if let LV::Boolean(b) = t.get::<LV>("clean_names")? {
            o.clean_names = b;
        }
        o.dates = string_list(&t, "dates")?;
        o.text = string_list(&t, "text")?;
        o.index = string_list(&t, "index")?;
    }
    Ok(o)
}

/// `xl.read_csv`, `xl.hash`, `xl.stat`, `xl.mkdir`, `xl.local_dir`: the small things a data-refresh
/// script needs around the programs it runs.
fn install_file_helpers(lua: &Lua) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;
    // xl.read_csv(path [, opts]) -> table of rows, column names first
    xl.set(
        "read_csv",
        lua.create_function(|lua, (path, opts): (String, Option<mlua::Table>)| {
            let o = csv_options(opts)?;
            let a = crate::csvio::read_array(Path::new(&path), &o, MAX_CELLS)
                .map_err(mlua::Error::runtime)?;
            to_lua(lua, &Value::Array(a))
        })?,
    )?;
    // xl.hash(text) -> SHA-256 as lowercase hex
    xl.set(
        "hash",
        lua.create_function(|_, text: mlua::LuaString| {
            use sha2::{Digest, Sha256};
            Ok(Sha256::digest(&*text.as_bytes())
                .iter()
                .map(|b| format!("{b:02x}"))
                .collect::<String>())
        })?,
    )?;
    // xl.stat(path) -> { size, mtime (seconds since 1970), is_dir, is_file } or nil
    xl.set(
        "stat",
        lua.create_function(|lua, path: String| {
            let Ok(m) = std::fs::metadata(&path) else {
                return Ok(LV::Nil);
            };
            let t = lua.create_table()?;
            t.set("size", m.len() as f64)?;
            t.set(
                "mtime",
                m.modified()
                    .ok()
                    .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
                    .map_or(0.0, |d| d.as_secs_f64()),
            )?;
            t.set("is_dir", m.is_dir())?;
            t.set("is_file", m.is_file())?;
            Ok(LV::Table(t))
        })?,
    )?;
    // xl.mkdir(path) -> true, or nil and a message
    xl.set(
        "mkdir",
        lua.create_function(|_, path: String| {
            Ok(match std::fs::create_dir_all(&path) {
                Ok(()) => (Some(true), None),
                Err(e) => (None, Some(format!("{path}: {e}"))),
            })
        })?,
    )?;
    // xl.local_dir() -> the machine folder (it may not exist yet)
    xl.set(
        "local_dir",
        lua.create_function(|_, ()| {
            Ok(crate::dbstore::local_dir().map(|p| p.to_string_lossy().into_owned()))
        })?,
    )
}

/// A database opened with `xl.db.open`.
struct LuaDb {
    db: Arc<crate::dbstore::Db>,
    env: Rc<SqlEnv>,
}

impl UserData for LuaDb {
    fn add_methods<M: mlua::UserDataMethods<Self>>(methods: &mut M) {
        // db:sql(sql [, params]) -> table of rows (column names first)
        methods.add_method(
            "sql",
            |lua, this, (sql, params): (String, Option<mlua::Table>)| {
                let (args, named) = split_params(params)?;
                let source = this.env.source();
                let out = this
                    .db
                    .run(&sql, &args, &named, source.as_ref(), this.env.limits())
                    .map_err(mlua::Error::runtime)?;
                to_lua(lua, &out.with_header())
            },
        );
        // db:records(sql [, params]) -> list of {column = value} tables
        methods.add_method(
            "records",
            |lua, this, (sql, params): (String, Option<mlua::Table>)| {
                let (args, named) = split_params(params)?;
                let source = this.env.source();
                let out = this
                    .db
                    .run(&sql, &args, &named, source.as_ref(), this.env.limits())
                    .map_err(mlua::Error::runtime)?;
                records_to_lua(lua, &out)
            },
        );
        // db:exec(script) -> runs statements that return no rows
        methods.add_method("exec", |_, this, script: String| {
            let source = this.env.source();
            this.db
                .exec(&script, source.as_ref(), &this.env.limits())
                .map_err(mlua::Error::runtime)
        });
        // db:import_csv(path, table [, opts]) -> { rows = n, columns = {...} }; replaces the table
        methods.add_method(
            "import_csv",
            |lua, this, (path, table, opts): (String, String, Option<mlua::Table>)| {
                let o = csv_options(opts)?;
                let done = this
                    .db
                    .import_csv(Path::new(&path), &table, &o, &this.env.limits())
                    .map_err(mlua::Error::runtime)?;
                let t = lua.create_table()?;
                t.set("rows", done.rows as i64)?;
                t.set("columns", lua.create_sequence_from(done.columns)?)?;
                Ok(t)
            },
        );
        methods.add_method("location", |_, this, ()| Ok(this.db.label()));
        methods.add_meta_method("__tostring", |_, this, ()| {
            Ok(format!("database({})", this.db.label()))
        });
    }
}

/// `xl.db.open(name [, {create = "embedded"|"local"|"folder", readonly = true}])` and `xl.db.list()`.
fn install_db_api(lua: &Lua, env: Rc<SqlEnv>) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;
    let db = lua.create_table()?;

    let e = Rc::clone(&env);
    db.set(
        "open",
        lua.create_function(move |_, (name, opts): (String, Option<mlua::Table>)| {
            use crate::dbstore::{self, Place};
            let (mut create, mut read_only) = (None::<String>, false);
            if let Some(o) = opts {
                if let LV::String(c) = o.get::<LV>("create")? {
                    create = Some(c.to_string_lossy());
                }
                read_only = matches!(o.get::<LV>("readonly")?, LV::Boolean(true));
            }
            let local = dbstore::local_dir();
            let found = dbstore::locate(&name, Some(&e.bundle), &e.workbook, local.as_deref());
            let (place, creating) = match (found, create) {
                (Ok(p), _) => (p, false),
                (Err(_), Some(where_)) => (
                    dbstore::creation_place(&name, &where_, &e.workbook, local.as_deref())
                        .map_err(mlua::Error::runtime)?,
                    true,
                ),
                (Err(m), None) => return Err(mlua::Error::runtime(m)),
            };
            let opened = match place {
                Place::Embedded(path) => {
                    let bytes = e.bundle.get(&path).map(|f| f.data.as_slice());
                    dbstore::open_embedded(&e.workbook, &path, bytes, creating, read_only)
                }
                Place::File(path) => dbstore::open_file(&path, creating, read_only).map(Arc::new),
            }
            .map_err(mlua::Error::runtime)?;
            Ok(LuaDb {
                db: opened,
                env: Rc::clone(&e),
            })
        })?,
    )?;

    let e = env;
    db.set(
        "list",
        lua.create_function(move |lua, ()| {
            let out = lua.create_table()?;
            let mut n = 0;
            let mut add = |name: String, place: &str| -> mlua::Result<()> {
                let t = lua.create_table()?;
                t.set("name", name)?;
                t.set("place", place)?;
                n += 1;
                out.raw_set(n, t)
            };
            for f in e.bundle.files() {
                if crate::vfs::is_database_file(&f.path) {
                    add(f.path.clone(), "embedded")?;
                }
            }
            let local = crate::dbstore::local_dir();
            for (dir, place) in [
                (e.workbook.parent().map(Path::to_path_buf), "folder"),
                (local, "local"),
            ] {
                let Some(dir) = dir.filter(|d| !d.as_os_str().is_empty()) else {
                    continue;
                };
                let Ok(rd) = std::fs::read_dir(dir) else {
                    continue;
                };
                let mut names: Vec<String> = rd
                    .filter_map(|e| e.ok())
                    .filter_map(|e| e.file_name().to_str().map(str::to_owned))
                    .filter(|n| crate::vfs::is_database_file(n))
                    .collect();
                names.sort();
                for name in names {
                    add(name, place)?;
                }
            }
            Ok(out)
        })?,
    )?;
    xl.set("db", db)
}

/// A formula region (see [`crate::regions`]) seen from Lua.
struct LuaRegion {
    api: Arc<dyn XlApi>,
    workbook: PathBuf,
    info: Rc<RefCell<crate::regions::RegionInfo>>,
}

fn region_call(api: &Arc<dyn XlApi>, workbook: &Path, op: xlapi::RegionOp) -> mlua::Result<Reply> {
    api.call(Op::Region {
        workbook: workbook.to_owned(),
        op,
    })
    .map_err(api_err)
}

fn region_of(api: &Arc<dyn XlApi>, workbook: &Path, reply: Reply) -> mlua::Result<LuaRegion> {
    match reply {
        Reply::Regions(mut v) if v.len() == 1 => Ok(LuaRegion {
            api: Arc::clone(api),
            workbook: workbook.to_owned(),
            info: Rc::new(RefCell::new(v.remove(0))),
        }),
        _ => Err(mlua::Error::runtime("unexpected reply from Excel")),
    }
}

fn bool_opt(opts: &Option<mlua::Table>, key: &str) -> mlua::Result<bool> {
    Ok(match opts {
        Some(t) => matches!(t.get::<LV>(key)?, LV::Boolean(true)),
        None => false,
    })
}

/// `{ old = "values" | "keep" | "clear" }`; the older `{ freeze = true/false }` still reads.
/// Without either, cells that leave a region are left alone.
fn old_opt(opts: &Option<mlua::Table>) -> mlua::Result<crate::regions::OldCells> {
    use crate::regions::OldCells;
    let Some(t) = opts else {
        return Ok(OldCells::Keep);
    };
    if let LV::String(w) = t.get::<LV>("old")? {
        let w = w.to_string_lossy();
        return OldCells::from_word(&w)
            .ok_or_else(|| mlua::Error::runtime("old must be 'values', 'keep' or 'clear'"));
    }
    Ok(match t.get::<LV>("freeze")? {
        LV::Boolean(true) => OldCells::Values,
        _ => OldCells::Keep,
    })
}

impl LuaRegion {
    fn name(&self) -> String {
        self.info.borrow().name.clone()
    }

    fn roll(&self, start: i32, end: i32, opts: &Option<mlua::Table>) -> mlua::Result<()> {
        let reply = region_call(
            &self.api,
            &self.workbook,
            xlapi::RegionOp::Roll {
                name: self.name(),
                start,
                end,
                old: old_opt(opts)?,
                overwrite: bool_opt(opts, "overwrite")?,
            },
        )?;
        if let Reply::Regions(mut v) = reply
            && v.len() == 1
        {
            *self.info.borrow_mut() = v.remove(0);
        }
        Ok(())
    }
}

impl UserData for LuaRegion {
    fn add_methods<M: mlua::UserDataMethods<Self>>(methods: &mut M) {
        methods.add_method("name", |_, this, ()| Ok(this.name()));
        methods.add_method("address", |_, this, ()| {
            Ok(this.info.borrow().address.clone())
        });
        methods.add_method("size", |_, this, ()| {
            let i = this.info.borrow();
            Ok((i.rows, i.cols))
        });
        // region:check() -> { ok, anchor, count, deviations = { {address, kind, text}, ... } }
        methods.add_method("check", |lua, this, ()| {
            let Reply::Report(r) = region_call(
                &this.api,
                &this.workbook,
                xlapi::RegionOp::Check { name: this.name() },
            )?
            else {
                return Err(mlua::Error::runtime("unexpected reply from Excel"));
            };
            let t = lua.create_table()?;
            t.set("ok", r.deviations.is_empty())?;
            t.set("anchor", r.anchor.as_str())?;
            t.set("count", r.deviations.len())?;
            let list = lua.create_table()?;
            for (i, d) in r.deviations.iter().enumerate() {
                let e = lua.create_table()?;
                e.set("address", d.address.as_str())?;
                e.set("kind", d.kind.word())?;
                e.set("text", d.text.as_str())?;
                list.raw_set(i + 1, e)?;
            }
            t.set("deviations", list)?;
            Ok(t)
        });
        // region:apply{ overwrite = false }: fill the whole region from its anchor again
        methods.add_method("apply", |_, this, opts: Option<mlua::Table>| {
            region_call(
                &this.api,
                &this.workbook,
                xlapi::RegionOp::Apply {
                    name: this.name(),
                    overwrite: bool_opt(&opts, "overwrite")?,
                },
            )
            .map(drop)
        });
        // Roll-forward. opts: { old = "values" | "keep" | "clear" } (or freeze = true) for the cells
        // that leave the region;
        // { overwrite = true } replaces cells that already hold something where the region grows.
        methods.add_method(
            "roll_start",
            |_, this, (n, opts): (Option<i32>, Option<mlua::Table>)| {
                this.roll(n.unwrap_or(1), 0, &opts)
            },
        );
        methods.add_method(
            "extend_end",
            |_, this, (n, opts): (Option<i32>, Option<mlua::Table>)| {
                this.roll(0, n.unwrap_or(1), &opts)
            },
        );
        methods.add_method(
            "roll",
            |_, this, (n, opts): (Option<i32>, Option<mlua::Table>)| {
                let n = n.unwrap_or(1);
                this.roll(n, n, &opts)
            },
        );
        // region:rule() -> "slide", "values"   /   region:set_rule("start" | "end" | "slide" | "none"
        // [, { old = "values" | "keep" | "clear" }]): how xl.region.roll_all moves this region and
        // what becomes of the cells that leave it.
        methods.add_method("rule", |_, this, ()| {
            let r = this.info.borrow().rule;
            Ok((r.mode.word().to_owned(), r.old.word().to_owned()))
        });
        methods.add_method(
            "set_rule",
            |_, this, (mode, opts): (String, Option<mlua::Table>)| {
                let mode = crate::regions::RollMode::from_word(&mode).ok_or_else(|| {
                    mlua::Error::runtime("the rule is 'none', 'start', 'end' or 'slide'")
                })?;
                // A rule's default is that old cells become values.
                let old = match &opts {
                    Some(t)
                        if !matches!(t.get::<LV>("old")?, LV::Nil)
                            || !matches!(t.get::<LV>("freeze")?, LV::Nil) =>
                    {
                        old_opt(&opts)?
                    }
                    _ => crate::regions::OldCells::Values,
                };
                let reply = region_call(
                    &this.api,
                    &this.workbook,
                    xlapi::RegionOp::SetRule {
                        name: this.name(),
                        rule: crate::regions::Rule { mode, old },
                    },
                )?;
                if let Reply::Regions(mut v) = reply
                    && v.len() == 1
                {
                    *this.info.borrow_mut() = v.remove(0);
                }
                Ok(())
            },
        );
        methods.add_method("remove", |_, this, ()| {
            region_call(
                &this.api,
                &this.workbook,
                xlapi::RegionOp::Remove { name: this.name() },
            )
            .map(drop)
        });
        methods.add_meta_method("__tostring", |_, this, ()| {
            let i = this.info.borrow();
            Ok(format!("region({} = {})", i.name, i.address))
        });
    }
}

/// `xl.region.define / get / at / list`
fn install_region_api(lua: &Lua, workbook: &Path, api: Option<Arc<dyn XlApi>>) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;
    let api = api.unwrap_or_else(|| Arc::new(Detached));
    let t = lua.create_table()?;
    let wb = workbook.to_owned();

    // xl.region.define("Forecast!E5:KJ5" [, { name = "dates", overwrite = false }])
    let (a, w) = (Arc::clone(&api), wb.clone());
    t.set(
        "define",
        lua.create_function(move |_, (address, opts): (String, Option<mlua::Table>)| {
            let label = match &opts {
                Some(o) => match o.get::<LV>("name")? {
                    LV::String(s) => Some(s.to_string_lossy()),
                    _ => None,
                },
                None => None,
            };
            let reply = region_call(
                &a,
                &w,
                xlapi::RegionOp::Define {
                    address,
                    label,
                    overwrite: bool_opt(&opts, "overwrite")?,
                },
            )?;
            region_of(&a, &w, reply)
        })?,
    )?;

    let (a, w) = (Arc::clone(&api), wb.clone());
    t.set(
        "list",
        lua.create_function(move |lua, ()| {
            let Reply::Regions(list) = region_call(&a, &w, xlapi::RegionOp::List)? else {
                return Err(mlua::Error::runtime("unexpected reply from Excel"));
            };
            let out = lua.create_table()?;
            for (i, info) in list.into_iter().enumerate() {
                out.raw_set(
                    i + 1,
                    LuaRegion {
                        api: Arc::clone(&a),
                        workbook: w.clone(),
                        info: Rc::new(RefCell::new(info)),
                    },
                )?;
            }
            Ok(out)
        })?,
    )?;

    let (a, w) = (Arc::clone(&api), wb.clone());
    t.set(
        "get",
        lua.create_function(move |_, name: String| {
            let Reply::Regions(list) = region_call(&a, &w, xlapi::RegionOp::List)? else {
                return Err(mlua::Error::runtime("unexpected reply from Excel"));
            };
            let info = list
                .into_iter()
                .find(|i| i.name == name)
                .ok_or_else(|| mlua::Error::runtime(format!("no formula region named '{name}'")))?;
            Ok(LuaRegion {
                api: Arc::clone(&a),
                workbook: w.clone(),
                info: Rc::new(RefCell::new(info)),
            })
        })?,
    )?;

    // xl.region.roll_all(n = 1 [, { overwrite = false }]) -> { rolled = { names }, failed = { {name, error} } }
    // Rolls every region that has a rule by n periods, each by its own rule.
    let (a, w) = (Arc::clone(&api), wb.clone());
    t.set(
        "roll_all",
        lua.create_function(move |lua, (n, opts): (Option<i32>, Option<mlua::Table>)| {
            let n = n.unwrap_or(1);
            let overwrite = bool_opt(&opts, "overwrite")?;
            let Reply::Regions(list) = region_call(&a, &w, xlapi::RegionOp::List)? else {
                return Err(mlua::Error::runtime("unexpected reply from Excel"));
            };
            let (rolled, failed) = (lua.create_table()?, lua.create_table()?);
            for info in list {
                let (start, end) = info.rule.steps(n);
                if info.rule.mode == crate::regions::RollMode::None {
                    continue;
                }
                let result = region_call(
                    &a,
                    &w,
                    xlapi::RegionOp::Roll {
                        name: info.name.clone(),
                        start,
                        end,
                        old: info.rule.old,
                        overwrite,
                    },
                );
                match result {
                    Ok(_) => rolled.raw_push(info.name.as_str())?,
                    Err(e) => {
                        let f = lua.create_table()?;
                        f.set("name", info.name.as_str())?;
                        f.set("error", e.to_string())?;
                        failed.raw_push(f)?;
                    }
                }
            }
            let out = lua.create_table()?;
            out.set("rolled", rolled)?;
            out.set("failed", failed)?;
            Ok(out)
        })?,
    )?;

    // xl.region.at("Forecast!G5") -> the region containing that cell, or nil
    let (a, w) = (api, wb);
    t.set(
        "at",
        lua.create_function(move |_, address: String| {
            let Reply::Regions(mut list) = region_call(&a, &w, xlapi::RegionOp::At { address })?
            else {
                return Err(mlua::Error::runtime("unexpected reply from Excel"));
            };
            Ok(list.pop().map(|info| LuaRegion {
                api: Arc::clone(&a),
                workbook: w.clone(),
                info: Rc::new(RefCell::new(info)),
            }))
        })?,
    )?;
    xl.set("region", t)
}

/// Stands in when no Excel is attached.
struct Detached;
impl XlApi for Detached {
    fn call(&self, _: Op) -> Result<Reply, xlapi::XlError> {
        Err(xlapi::XlError::new("Excel is not available here"))
    }
}

fn strip_bom(data: &[u8]) -> &[u8] {
    data.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(data)
}

const NATIVE_EXTS: [&str; 3] = [".dll", ".so", ".dylib"];

/// Writes the native libraries of the tree under `<cache>/<digest prefix>/` and returns that
/// directory (or `None` if there are none). Existing files of the right size are reused.
fn extract_native(bundle: &Bundle, cache: &Path) -> std::io::Result<Option<PathBuf>> {
    let natives: Vec<_> = bundle
        .files()
        .iter()
        .filter(|f| {
            NATIVE_EXTS
                .iter()
                .any(|e| f.path.to_lowercase().ends_with(e))
        })
        .collect();
    if natives.is_empty() {
        return Ok(None);
    }
    let root = cache.join(&bundle.digest()[..16]);
    for f in natives {
        let target = root.join(&f.path); // paths are validated by `Bundle`: no `..`, no drive letters
        if std::fs::metadata(&target).is_ok_and(|m| m.len() == f.data.len() as u64) {
            continue;
        }
        if let Some(parent) = target.parent() {
            std::fs::create_dir_all(parent)?;
        }
        std::fs::write(&target, &f.data)?;
    }
    Ok(Some(root))
}

/// `xl.read_file(path)`, `xl.files()` and a `require` searcher backed by the embedded tree.
fn install_embedded_fs(lua: &Lua, bundle: &Rc<Bundle>) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;

    let b = Rc::clone(bundle);
    xl.set(
        "read_file",
        lua.create_function(move |lua, path: String| match b.get(&path) {
            Some(f) => Ok(LV::String(lua.create_string(&f.data)?)),
            None => Ok(LV::Nil),
        })?,
    )?;

    let b = Rc::clone(bundle);
    xl.set(
        "files",
        lua.create_function(move |lua, ()| {
            let t = lua.create_table()?;
            for (i, f) in b.files().iter().enumerate() {
                t.raw_set(i + 1, f.path.as_str())?;
            }
            Ok(t)
        })?,
    )?;

    let b = Rc::clone(bundle);
    let searcher = lua.create_function(move |lua, name: String| -> mlua::Result<MultiValue> {
        let rel = name.replace('.', "/");
        let candidates = [format!("{rel}.lua"), format!("{rel}/init.lua")];
        for path in &candidates {
            if let Some(f) = b.get(path) {
                let loader = lua
                    .load(strip_bom(&f.data))
                    .set_name(format!("={}", f.path))
                    .set_mode(ChunkMode::Text)
                    .into_function()?;
                return Ok(MultiValue::from_iter([
                    LV::Function(loader),
                    LV::String(lua.create_string(&f.path)?),
                ]));
            }
        }
        let tried: String = candidates
            .iter()
            .map(|c| format!("\n\tno embedded file '{c}'"))
            .collect();
        Ok(MultiValue::from_iter([LV::String(
            lua.create_string(&tried)?,
        )]))
    })?;
    let package: mlua::Table = lua.globals().get("package")?;
    let searchers: mlua::Table = package.get("searchers")?;
    let insert: Function = lua.globals().get::<mlua::Table>("table")?.get("insert")?;
    insert.call::<()>((searchers, 2, searcher))
}

/// `xl.run(program | {program, args..., cwd=, timeout=, stdin=}, [opts])` and
/// `xl.shell(line, [opts])`: run an external program with captured output and a real timeout.
/// Returns `{code, stdout, stderr, ok, timed_out, truncated}`, or `nil, message` if it could not
/// be started. No console window appears.
fn install_process_api(lua: &Lua, limits: Rc<RefCell<Option<Limits>>>) -> mlua::Result<()> {
    let xl: mlua::Table = lua.globals().get("xl")?;
    // A background call that is cancelled also kills the program it is waiting for.
    let cancel_of = move || limits.borrow().as_ref().map(|l| Arc::clone(&l.cancel));
    let c = cancel_of.clone();
    xl.set(
        "run",
        lua.create_function(move |lua, (spec, opts): (LV, Option<mlua::Table>)| {
            run_from_lua(lua, spec, opts, false, c())
        })?,
    )?;
    xl.set(
        "shell",
        lua.create_function(move |lua, (line, opts): (String, Option<mlua::Table>)| {
            run_from_lua(
                lua,
                LV::String(lua.create_string(&line)?),
                opts,
                true,
                cancel_of(),
            )
        })?,
    )
}

fn word(v: &LV) -> Option<String> {
    match v {
        LV::String(s) => Some(s.to_string_lossy()),
        LV::Integer(i) => Some(i.to_string()),
        LV::Number(n) => Some(n.to_string()),
        _ => None,
    }
}

fn run_from_lua(
    lua: &Lua,
    spec: LV,
    opts: Option<mlua::Table>,
    shell: bool,
    cancel: Option<Arc<AtomicBool>>,
) -> mlua::Result<MultiValue> {
    let mut run = match (&spec, shell) {
        (LV::String(line), true) => RunSpec::shell(&line.to_string_lossy()),
        (LV::String(program), false) => RunSpec::new(program.to_string_lossy(), Vec::new()),
        (LV::Table(t), false) => {
            let mut words: Vec<String> = Vec::new();
            for i in 1..=t.raw_len() {
                let v: LV = t.raw_get(i)?;
                words.push(word(&v).ok_or_else(|| {
                    mlua::Error::runtime("xl.run: arguments must be strings or numbers")
                })?);
            }
            if words.is_empty() {
                return Err(mlua::Error::runtime("xl.run: missing program"));
            }
            let program = words.remove(0);
            RunSpec::new(program, words)
        }
        _ => {
            return Err(mlua::Error::runtime(
                "xl.run: expected a program name or a list {program, args...}",
            ));
        }
    };
    // Options live in the spec table itself, or in the second argument.
    let options = opts.or_else(|| match &spec {
        LV::Table(t) => Some(t.clone()),
        _ => None,
    });
    if let Some(o) = options {
        if let LV::String(cwd) = o.get::<LV>("cwd")? {
            run.cwd = Some(PathBuf::from(cwd.to_string_lossy()));
        }
        match o.get::<LV>("timeout")? {
            LV::Integer(n) => run.timeout = Duration::from_secs_f64((n as f64).max(0.01)),
            LV::Number(n) => run.timeout = Duration::from_secs_f64(n.max(0.01)),
            _ => {}
        }
        if let LV::String(input) = o.get::<LV>("stdin")? {
            run.stdin = Some(input.as_bytes().to_vec());
        }
        // window = true: a visible console, for programs that sign in with a dialog.
        run.show_console = matches!(o.get::<LV>("window")?, LV::Boolean(true));
    }
    run.cancel = cancel;
    match procrun::run(&run) {
        Err(e) => Ok(MultiValue::from_iter([
            LV::Nil,
            LV::String(lua.create_string(format!("cannot run '{}': {e}", run.program))?),
        ])),
        Ok(r) if r.cancelled => Err(mlua::Error::runtime("cancelled")),
        Ok(r) => {
            let t = lua.create_table()?;
            match r.code {
                Some(c) => t.set("code", c)?,
                None => t.set("code", LV::Nil)?,
            }
            t.set("stdout", lua.create_string(&r.stdout)?)?;
            t.set("stderr", lua.create_string(&r.stderr)?)?;
            t.set("ok", r.ok())?;
            t.set("timed_out", r.timed_out)?;
            t.set("truncated", r.truncated)?;
            Ok(MultiValue::from_iter([LV::Table(t)]))
        }
    }
}

/// `require` finds modules next to the workbook, then in `%APPDATA%\xll-bridge\local`.
fn configure_package_path(lua: &Lua, workbook: &Path, native: Option<&Path>) -> mlua::Result<()> {
    let mut dirs: Vec<PathBuf> = Vec::new();
    if let Some(d) = workbook.parent() {
        dirs.push(d.to_path_buf());
    }
    // The machine folder: Lua modules, native modules and databases shared by every workbook.
    if let Some(p) = crate::dbstore::local_dir() {
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
    let mut c_paths: Vec<String> = dirs
        .iter()
        .map(|d| format!("{}{sep}?.{dll_ext}", d.to_string_lossy()))
        .collect();
    // Embedded native modules go first: `require("a.b")` finds `<cache>/a/b.dll`.
    if let Some(n) = native {
        c_paths.insert(0, format!("{}{sep}?.{dll_ext}", n.to_string_lossy()));
    }

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
    use crate::package::testutil::bundle_of;

    fn host(code: &str) -> LuaHost {
        host_of(&bundle_of(&[("t.lua", code)]), None)
    }

    fn host_of(b: &Bundle, cache: Option<&Path>) -> LuaHost {
        LuaHost::new(
            b,
            &HostContext {
                workbook: Path::new("/tmp/book.xlsx"),
                cache_dir: cache,
                xl: None,
            },
        )
        .unwrap()
    }

    fn grid_host(code: &str, grid: Arc<crate::xlapi::fake::MemGrid>) -> LuaHost {
        LuaHost::new(
            &bundle_of(&[("t.lua", code)]),
            &HostContext {
                workbook: Path::new("/tmp/book.xlsx"),
                cache_dir: None,
                xl: Some(grid),
            },
        )
        .unwrap()
    }

    #[test]
    fn ranges_read_and_write() {
        use crate::xlapi::fake::MemGrid;
        let g = MemGrid::new();
        g.put("Data", "A1", Value::Num(1.0));
        g.put("Data", "B1", Value::Num(2.0));
        g.put("Data", "A2", Value::Str("x".into()));
        g.put("Other", "A1", Value::Num(9.0));
        let h = grid_host(
            r#"
            function one() return xl.range("A1"):value() end
            function other() return xl.range("Other!A1"):value() end
            function sum()
              local t, s = xl.range("A1:B1"):values(), 0
              for _, v in ipairs(t[1]) do s = s + v end
              return s
            end
            function put() xl.range("D1"):set({{1, 2}, {3, 4}}) return xl.range("D1:E2"):values() end
            function fill() xl.range("G1:G3"):set(7) return xl.range("G1:G3"):values() end
            function col() xl.range("H1"):set({5, 6, 7}) return xl.range("H1:H3"):values() end
            function dims() local r, c = xl.range("A1:C4"):size() return r * 10 + c end
            function sheets() return xl.sheets() end
            function addr() return xl.range("Other!B2:C3"):address() end
            function bad() return xl.range("Nope!A1") end
            "#,
            Arc::clone(&g),
        );
        assert_eq!(h.call("one", &[]).unwrap(), Value::Num(1.0));
        assert_eq!(h.call("other", &[]).unwrap(), Value::Num(9.0));
        assert_eq!(h.call("sum", &[]).unwrap(), Value::Num(3.0));
        assert!(
            matches!(h.call("put", &[]).unwrap(), Value::Array(a) if a.rows == 2 && a.cols == 2)
        );
        assert_eq!(g.get("Data", "E2"), Value::Num(4.0));
        h.call("fill", &[]).unwrap();
        assert_eq!(g.get("Data", "G3"), Value::Num(7.0));
        h.call("col", &[]).unwrap();
        assert_eq!(g.get("Data", "H3"), Value::Num(7.0));
        assert_eq!(h.call("dims", &[]).unwrap(), Value::Num(43.0));
        assert!(matches!(h.call("sheets", &[]).unwrap(), Value::Array(a) if a.rows == 2));
        assert_eq!(
            h.call("addr", &[]).unwrap(),
            Value::Str("Other!$B$2:$C$3".into())
        );
        let e = h.call("bad", &[]).unwrap_err().to_string();
        assert!(e.contains("no sheet named"), "{e}");
    }

    #[test]
    fn ranges_are_read_only_for_lua_call_and_absent_without_excel() {
        use crate::xlapi::{ReadOnly, fake::MemGrid};
        let g = MemGrid::new();
        let h = LuaHost::new(
            &bundle_of(&[("t.lua", "function w() xl.range('A1'):set(1) end function r() return xl.range('A1'):value() end")]),
            &HostContext {
                workbook: Path::new("/tmp/book.xlsx"),
                cache_dir: None,
                xl: Some(Arc::new(ReadOnly(g))),
            },
        )
        .unwrap();
        assert_eq!(h.call("r", &[]).unwrap(), Value::Empty);
        assert!(
            h.call("w", &[])
                .unwrap_err()
                .to_string()
                .contains("LUA.ASYNC")
        );
        let h = host("function r() return xl.range('A1'):value() end");
        assert!(
            h.call("r", &[])
                .unwrap_err()
                .to_string()
                .contains("not available")
        );
    }

    #[test]
    fn macros_are_listed_and_run_with_limits() {
        let b = bundle_of(&[(
            "t.lua",
            "xl.macro('zeta', function() end) xl.macro('alpha', function() xl.range('A1'):set(5) end) \
             xl.macro('boom', function() error('kaput') end)",
        )]);
        let g = crate::xlapi::fake::MemGrid::new();
        let h = LuaHost::new_worker(
            &b,
            &HostContext {
                workbook: Path::new("/tmp/book.xlsx"),
                cache_dir: None,
                xl: Some(Arc::clone(&g) as Arc<dyn XlApi>),
            },
        )
        .unwrap();
        assert_eq!(h.macro_names(), ["alpha", "boom", "zeta"]);
        let limits = Limits {
            deadline: Instant::now() + Duration::from_secs(5),
            cancel: Arc::new(AtomicBool::new(false)),
        };
        h.call_macro("alpha", &limits).unwrap();
        assert_eq!(g.get("Data", "A1"), Value::Num(5.0));
        assert!(
            h.call_macro("boom", &limits)
                .unwrap_err()
                .to_string()
                .contains("kaput")
        );
        assert!(matches!(
            h.call_macro("nope", &limits),
            Err(HostError::Undefined(_))
        ));
    }

    #[test]
    fn sql_from_lua_over_cells_tables_and_dates() {
        use crate::xlapi::fake::MemGrid;
        let g = MemGrid::new();
        for (a, v) in [
            ("A1", Value::Str("name".into())),
            ("B1", Value::Str("qty".into())),
            ("A2", Value::Str("bolt".into())),
            ("B2", Value::Num(10.0)),
            ("A3", Value::Str("nut".into())),
            ("B3", Value::Num(25.0)),
        ] {
            g.put("Data", a, v);
        }
        g.tables
            .lock()
            .unwrap()
            .insert("Orders".into(), "Data!A1:B3".into());
        let h = grid_host(
            r#"
            function rows() return xl.sql("select name from [Orders] where qty > ?", {15}) end
            function named()
              local r = xl.sql_records("select name, qty from [Data!A1:B3] where qty >= :m order by qty", {m = 10})
              return #r .. ":" .. r[1].name .. ":" .. r[2].qty
            end
            function chain()
              xl.sql_exec("create temp table t(a); insert into t values (1),(2)")
              return 1
            end
            function wsf() return xl.query("select count(*) as n from t1 where qty > ?2", xl.range("A1:B3"):values(), 5) end
            function bad() return xl.sql("select * from [Nope]") end
            function dates() return xl.date_to_iso(45351) .. " " .. xl.iso_to_date("2024-02-29") end
            "#,
            Arc::clone(&g),
        );
        let Value::Array(a) = h.call("rows", &[]).unwrap() else {
            panic!()
        };
        assert_eq!(
            a.data,
            vec![Value::Str("name".into()), Value::Str("nut".into())]
        );
        assert_eq!(
            h.call("named", &[]).unwrap(),
            Value::Str("2:bolt:25.0".into())
        );
        assert_eq!(h.call("chain", &[]).unwrap(), Value::Num(1.0));
        let Value::Array(w) = h.call("wsf", &[]).unwrap() else {
            panic!()
        };
        assert_eq!(w.data[1], Value::Num(2.0));
        assert!(h.call("bad", &[]).unwrap_err().to_string().contains("Nope"));
        assert_eq!(
            h.call("dates", &[]).unwrap(),
            Value::Str("2024-02-29 45351.0".into())
        );
    }

    #[test]
    fn databases_embedded_and_beside_the_workbook() {
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("dbbook.xlsx");
        let h = LuaHost::new(
            &bundle_of(&[(
                "t.lua",
                r#"
                function make()
                  local d = xl.db.open("notes", {create = "embedded"})
                  d:exec("create table n(t, v)")
                  d:sql("insert into n values (?, ?)", {"hi", 5})
                  return d:location()
                end
                function read()
                  local d = xl.db.open("notes")           -- found again, with its data
                  return d:sql("select t, v from n")[2][1]
                end
                function records()
                  local r = xl.db.open("notes"):records("select t, v from n")
                  return r[1].t .. r[1].v
                end
                function attached()
                  return xl.sql("select v from [db:notes].n")[2][1]
                end
                function beside()
                  local d = xl.db.open("side", {create = "folder"})
                  d:exec("create table s(x); insert into s values (3)")
                  return xl.db.open("side"):sql("select x from s")[2][1]
                end
                function missing() return xl.db.open("nothing") end
                function readonly() return xl.db.open("notes", {readonly = true}):exec("drop table n") end
                function listed()
                  local names = {}
                  for _, d in ipairs(xl.db.list()) do names[#names + 1] = d.place .. ":" .. d.name end
                  return table.concat(names, ",")
                end
                "#,
            )]),
            &HostContext {
                workbook: &wb,
                cache_dir: None,
                xl: None,
            },
        )
        .unwrap();
        assert_eq!(
            h.call("make", &[]).unwrap(),
            Value::Str("embedded:notes.db".into())
        );
        assert_eq!(h.call("read", &[]).unwrap(), Value::Str("hi".into()));
        assert_eq!(h.call("records", &[]).unwrap(), Value::Str("hi5.0".into()));
        assert_eq!(h.call("attached", &[]).unwrap(), Value::Num(5.0));
        assert_eq!(h.call("beside", &[]).unwrap(), Value::Num(3.0));
        assert!(dir.path().join("side.db").is_file());
        let e = h.call("missing", &[]).unwrap_err().to_string();
        assert!(e.contains("not found"), "{e}");
        assert!(
            h.call("readonly", &[])
                .unwrap_err()
                .to_string()
                .contains("read-only")
        );
        assert_eq!(
            h.call("listed", &[]).unwrap(),
            Value::Str("folder:side.db".into())
        );
        crate::dbstore::forget_workbook(&wb);
    }

    #[test]
    fn csv_files_load_into_databases_and_helpers_for_refresh_scripts() {
        use crate::xlapi::fake::MemGrid;
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("csvbook.xlsx");
        let csv = dir.path().join("gl.csv");
        std::fs::write(
            &csv,
            "'Dim'[BU],[Month],[Amount]\nNorth,2024-01-31,5\nSouth,2024-02-29,7.5\nNorth,2024-02-29,1\n",
        )
        .unwrap();
        let g = MemGrid::new();
        let h = LuaHost::new(
            &bundle_of(&[(
                "t.lua",
                r#"
                function load(csv)
                  local d = xl.db.open("gl", {create = "folder"})
                  local r = d:import_csv(csv, "gl", {dates = {"Month"}, index = {"BU"}})
                  local t = d:sql("select BU, sum(Amount) from gl group by BU order by BU")
                  return r.rows .. ":" .. r.columns[2] .. ":" .. t[2][1] .. "=" .. t[2][2]
                end
                function read(csv)
                  local t = xl.read_csv(csv, {dates = {"Month"}})
                  return #t .. ":" .. t[1][3] .. ":" .. t[2][2]
                end
                function helpers(path)
                  local s = xl.stat(path)
                  return xl.hash("abc"):sub(1, 8) .. ":" .. tostring(s.is_file) .. ":" .. tostring(xl.stat(path .. "x"))
                end
                function table_of()
                  xl.range("A1"):set({{"k", "v"}, {1, 2}})
                  xl.range("A1:B2"):as_table("Report")
                  return xl.sql("select v from [Report]")[2][1]
                end
                "#,
            )]),
            &HostContext {
                workbook: &wb,
                cache_dir: None,
                xl: Some(Arc::clone(&g) as Arc<dyn XlApi>),
            },
        )
        .unwrap();
        let path = Value::Str(csv.to_string_lossy().into_owned());
        assert_eq!(
            h.call("load", std::slice::from_ref(&path)).unwrap(),
            Value::Str("3:Month:North=6.0".into())
        );
        assert_eq!(
            h.call("read", std::slice::from_ref(&path)).unwrap(),
            Value::Str("4:Amount:45322.0".into())
        );
        assert_eq!(
            h.call("helpers", &[path]).unwrap(),
            Value::Str("ba7816bf:true:nil".into())
        );
        // A range becomes a named table that SQL (and, in Excel, PivotTables) can use.
        assert_eq!(h.call("table_of", &[]).unwrap(), Value::Num(2.0));
        assert!(g.tables.lock().unwrap().contains_key("Report"));
        assert!(dir.path().join("gl.db").is_file());
        crate::dbstore::forget_workbook(&wb);
    }

    #[test]
    fn the_fabric_example_caches_pulls_on_disk() {
        let dir = tempfile::tempdir().unwrap();
        let (cache, src) = (dir.path().join("local"), dir.path().join("src.csv"));
        std::fs::write(&src, "'Dim'[BU],[Amount]\nNorth,5\nSouth,7\n").unwrap();
        // LOCAL and SRC are plain globals for the script, set by a file that loads first.
        let code = format!(
            "LOCAL = [[{}]] SRC = [[{}]]\n",
            cache.display(),
            src.display()
        );
        let bundle = bundle_of(&[
            ("fabric.lua", include_str!("../examples/fabric.lua")),
            ("a_globals.lua", &code),
            (
                "t.lua",
                r#"
                local fabric = require("fabric")
                xl.local_dir = function() return LOCAL end
                pulls = 0
                fabric.config.command = function(c, dax_file, out_file)
                  pulls = pulls + 1
                  if package.config:sub(1, 1) == "\\" then
                    return { "cmd", "/C", "copy", SRC, out_file }
                  end
                  return { "cp", SRC, out_file }
                end
                fabric.config.refresh_after = "00:00"
                function run(dax, force)
                  local t = fabric.table(dax, { force = force })
                  return pulls .. ":" .. #t .. ":" .. t[1][1] .. ":" .. t[3][2]
                end
                "#,
            ),
        ]);
        let h = LuaHost::new(
            &bundle,
            &HostContext {
                workbook: &dir.path().join("b.xlsx"),
                cache_dir: None,
                xl: None,
            },
        )
        .unwrap();
        let q = |dax: &str, force: bool| {
            h.call("run", &[Value::Str(dax.into()), Value::Bool(force)])
                .unwrap()
        };
        // The first call pulls; the same DAX again is served from disk; other DAX pulls again.
        assert_eq!(q("EVALUATE A", false), Value::Str("1:3:BU:7.0".into()));
        assert_eq!(q("EVALUATE A", false), Value::Str("1:3:BU:7.0".into()));
        assert_eq!(q("EVALUATE B", false), Value::Str("2:3:BU:7.0".into()));
        assert_eq!(q("EVALUATE A", true), Value::Str("3:3:BU:7.0".into()));
        assert!(cache.join("fabric-cache").is_dir());
    }

    #[test]
    fn region_rules_drive_a_workbook_roll() {
        use crate::xlapi::fake::MemGrid;
        let g = MemGrid::new();
        {
            let mut s = g.sheets.lock().unwrap();
            let d = s.get_mut("Data").unwrap();
            for c in [1, 2, 3, 6, 7, 8] {
                d.formulas.insert((0, c), "=RC[-1]+1".into());
            }
            // a second row of actuals and a fixed block
            for c in [1, 2] {
                d.formulas.insert((2, c), "=RC[-1]+1".into());
                d.formulas.insert((4, c), "=RC[-1]+1".into());
            }
        }
        let h = grid_host(
            r#"
            function setup()
              xl.region.define("B1:D1", { name = "fc" }):set_rule("start")
              xl.region.define("G1:I1", { name = "roll" }):set_rule("slide", { old = "clear" })
              xl.region.define("B3:C3", { name = "act" }):set_rule("end")
              xl.region.define("B5:C5", { name = "fixed" })
            end
            function rules()
              local out = {}
              for _, r in ipairs(xl.region.list()) do
                local m, o = r:rule()
                out[#out + 1] = r:name() .. "=" .. m .. "/" .. o
              end
              table.sort(out)
              return table.concat(out, " ")
            end
            function go()
              local res = xl.region.roll_all(1)
              local out = {}
              for _, n in ipairs(res.rolled) do out[#out + 1] = n end
              table.sort(out)
              return table.concat(out, ",") .. " failed=" .. #res.failed
            end
            function where_are()
              local out = {}
              for _, r in ipairs(xl.region.list()) do out[#out + 1] = r:name() .. "@" .. r:address() end
              table.sort(out)
              return table.concat(out, " ")
            end
            function redefine()
              local r = xl.region.define("B5:D5")
              local m, o = r:rule()
              return r:name() .. "@" .. r:address() .. " " .. m
            end
            function extend_left() return pcall(function() xl.region.define("A5:D5") end) end
            function shrink() return xl.region.define("B5:C5"):address() end
            function several() return pcall(function() xl.region.define("A3:E5") end) end
            function bad() return pcall(function() xl.region.get("fc"):set_rule("sideways") end) end
            "#,
            Arc::clone(&g),
        );
        h.call("setup", &[]).unwrap();
        assert_eq!(
            h.call("rules", &[]).unwrap(),
            Value::Str("act=end/values fc=start/values fixed=none/values roll=slide/clear".into())
        );
        assert_eq!(
            h.call("go", &[]).unwrap(),
            Value::Str("act,fc,roll failed=0".into())
        );
        assert_eq!(
            h.call("where_are", &[]).unwrap(),
            Value::Str("act@Data!B3:D3 fc@Data!C1:D1 fixed@Data!B5:C5 roll@Data!H1:J1".into())
        );
        assert_eq!(h.call("bad", &[]).unwrap(), Value::Bool(false));
        // Defining again from the same anchor extends that region and keeps its name and rule.
        assert_eq!(
            h.call("redefine", &[]).unwrap(),
            Value::Str("fixed@Data!B5:D5 none".into())
        );
        // The anchor never moves: a range starting elsewhere that overlaps a region is refused,
        // while shrinking from the same anchor redefines it.
        assert_eq!(h.call("extend_left", &[]).unwrap(), Value::Bool(false));
        assert_eq!(
            h.call("shrink", &[]).unwrap(),
            Value::Str("Data!B5:C5".into())
        );
        assert_eq!(h.call("several", &[]).unwrap(), Value::Bool(false));
    }

    #[test]
    fn formula_regions_from_lua() {
        use crate::xlapi::fake::MemGrid;
        let g = MemGrid::new();
        // A row of dates: seed in A1, the formula in B1 only.
        g.sheets
            .lock()
            .unwrap()
            .get_mut("Data")
            .unwrap()
            .formulas
            .insert((0, 1), "=EOMONTH(RC[-1],1)".into());
        let h = grid_host(
            r#"
            function setup()
              local r = xl.region.define("B1:F1", { name = "dates" })
              return r:name() .. " " .. r:address()
            end
            function clean() return xl.region.get("dates"):check().ok end
            function spoil() xl.range("D1"):set(99) xl.range("E1"):set_formulas({{"=RC[-2]"}}) end
            function report()
              local c = xl.region.get("dates"):check()
              local out = {}
              for _, d in ipairs(c.deviations) do out[#out + 1] = d.address .. ":" .. d.kind end
              return c.count .. " " .. table.concat(out, ",")
            end
            function refused() return pcall(function() xl.region.get("dates"):apply() end) end
            function repair() xl.region.get("dates"):apply { overwrite = true } return xl.region.get("dates"):check().ok end
            function where_is() local r = xl.region.at("C1") return r and r:name() or "none" end
            function outside() return xl.region.at("A5") == nil end
            function slide() local r = xl.region.get("dates") r:roll(2) return r:address() end
            function grow() local r = xl.region.get("dates") r:extend_end(3) return r:address() end
            function forecast() local r = xl.region.get("dates") r:roll_start(1) return r:address() end
            function gone() xl.region.get("dates"):remove() return #xl.region.list() end
            "#,
            Arc::clone(&g),
        );
        assert_eq!(
            h.call("setup", &[]).unwrap(),
            Value::Str("dates Data!B1:F1".into())
        );
        assert_eq!(h.call("clean", &[]).unwrap(), Value::Bool(true));
        h.call("spoil", &[]).unwrap();
        assert_eq!(
            h.call("report", &[]).unwrap(),
            Value::Str("2 D1:constant,E1:formula".into())
        );
        // Without overwrite the apply is refused and says what is in the way.
        assert_eq!(h.call("refused", &[]).unwrap(), Value::Bool(false));
        assert_eq!(h.call("repair", &[]).unwrap(), Value::Bool(true));
        assert_eq!(h.call("where_is", &[]).unwrap(), Value::Str("dates".into()));
        assert_eq!(h.call("outside", &[]).unwrap(), Value::Bool(true));
        // Rolling: slide the window by two, grow the end by three, move the start by one.
        assert_eq!(
            h.call("slide", &[]).unwrap(),
            Value::Str("Data!D1:H1".into())
        );
        assert_eq!(
            h.call("grow", &[]).unwrap(),
            Value::Str("Data!D1:K1".into())
        );
        assert_eq!(
            h.call("forecast", &[]).unwrap(),
            Value::Str("Data!E1:K1".into())
        );
        assert_eq!(h.call("gone", &[]).unwrap(), Value::Num(0.0));
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
    fn sandbox_has_no_ambient_authority() {
        let h = LuaHost::sandboxed().unwrap();
        for name in [
            "io", "os", "package", "debug", "dofile", "loadfile", "load", "require",
        ] {
            let code = format!("return {name} == nil");
            assert_eq!(
                h.eval(&code, &[]).unwrap(),
                Value::Bool(true),
                "{name} must be absent"
            );
        }
        assert_eq!(h.eval("1 +2", &[]).unwrap(), Value::Num(3.0));
        assert_eq!(
            h.eval("string.rep('ab', 3)", &[]).unwrap(),
            Value::Str("ababab".into())
        );
        assert_eq!(h.eval("xl.NA", &[]).unwrap(), Value::Err(ErrorCode::NA));
    }

    #[test]
    fn sandbox_stops_runaway_code_and_recovers() {
        let h = LuaHost::sandboxed().unwrap();
        let Err(HostError::Lua(m)) = h.eval("while true do end", &[]) else {
            panic!("loop must be aborted")
        };
        assert!(m.contains("instruction limit"), "{m}");
        // The budget is per call, so the state stays usable.
        assert_eq!(h.eval("2 * 21", &[]).unwrap(), Value::Num(42.0));
    }

    #[test]
    fn sandbox_limits_memory() {
        let h = LuaHost::sandboxed().unwrap();
        assert!(h.eval("string.rep('x', 200 * 1024 * 1024)", &[]).is_err());
    }

    #[test]
    fn binary_chunks_are_rejected() {
        let h = LuaHost::sandboxed().unwrap();
        assert!(h.eval("\x1bLua", &[]).is_err());
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

    #[test]
    fn require_reads_modules_from_the_embedded_tree() {
        let b = bundle_of(&[
            (
                "main.lua",
                "local u = require('lib.util') function twice(x) return u.double(x) end function ver() return require('pkg').name end",
            ),
            (
                "lib/util.lua",
                "return { double = function(x) return x * 2 end }",
            ),
            ("pkg/init.lua", "return { name = 'pkg-init' }"),
        ]);
        let h = host_of(&b, None);
        assert_eq!(
            h.call("twice", &[Value::Num(4.0)]).unwrap(),
            Value::Num(8.0)
        );
        assert_eq!(h.call("ver", &[]).unwrap(), Value::Str("pkg-init".into()));
        let Value::Str(m) = h
            .eval("select(2, pcall(require, 'nope.missing'))", &[])
            .unwrap()
        else {
            panic!()
        };
        assert!(m.contains("no embedded file 'nope/missing.lua'"), "{m}");
    }

    #[test]
    fn only_root_lua_files_run_at_load_and_in_name_order() {
        let b = bundle_of(&[
            ("b.lua", "order = (order or '') .. 'b'"),
            ("a.lua", "order = (order or '') .. 'a'"),
            ("lib/never.lua", "order = 'WRONG'"),
            ("notes.txt", "order = 'WRONG'"),
        ]);
        let h = host_of(&b, None);
        assert_eq!(h.eval("order", &[]).unwrap(), Value::Str("ab".into()));
    }

    #[test]
    fn read_file_and_files_expose_data_binary_safely() {
        let mut b = bundle_of(&[("main.lua", "")]);
        b.put_file("data/blob.bin", vec![0, 255, 10, 0], true)
            .unwrap();
        let h = host_of(&b, None);
        assert_eq!(
            h.eval("#xl.read_file('DATA/blob.bin')", &[]).unwrap(),
            Value::Num(4.0)
        );
        assert_eq!(
            h.eval("xl.read_file('missing')", &[]).unwrap(),
            Value::Empty
        );
        assert_eq!(h.eval("#xl.files()", &[]).unwrap(), Value::Num(2.0));
    }

    #[test]
    fn native_libraries_are_extracted_once_into_the_cache() {
        let cache = tempfile::tempdir().unwrap();
        let mut b = bundle_of(&[("main.lua", "")]);
        b.put_file("native/thing.dll", vec![1, 2, 3], true).unwrap();
        b.put_file("data.bin", vec![9], true).unwrap();
        let _ = host_of(&b, Some(cache.path()));
        let dir = cache.path().join(&b.digest()[..16]);
        assert_eq!(
            std::fs::read(dir.join("native/thing.dll")).unwrap(),
            [1, 2, 3]
        );
        assert!(
            !dir.join("data.bin").exists(),
            "only native libraries are extracted"
        );
        // Reused on the next activation, and cpath points at it.
        let h = host_of(&b, Some(cache.path()));
        let Value::Str(cpath) = h.eval("package.cpath", &[]).unwrap() else {
            panic!()
        };
        assert!(cpath.contains(&b.digest()[..16]));
    }

    #[test]
    fn bom_prefixed_scripts_load() {
        let mut b = Bundle::new();
        b.put_file("main.lua", b"\xEF\xBB\xBFx = 41".to_vec(), true)
            .unwrap();
        assert_eq!(
            host_of(&b, None).eval("x + 1", &[]).unwrap(),
            Value::Num(42.0)
        );
    }

    fn quick_shell(line: &str) -> String {
        line.to_owned()
    }

    #[test]
    fn lua_can_run_programs_with_captured_output_and_a_timeout() {
        let h = host(
            "function out() local r = xl.shell('echo hi') return (r.stdout:gsub('%s+$', '')) end
             function missing() local r, err = xl.run{'definitely-not-a-real-program-xyz'} return r == nil and err ~= nil end
             function code() return xl.shell('exit 3').code end
             function slow(line) local r = xl.shell(line, { timeout = 0.3 }) return r.timed_out end",
        );
        assert_eq!(h.call("out", &[]).unwrap(), Value::Str("hi".into()));
        assert_eq!(h.call("missing", &[]).unwrap(), Value::Bool(true));
        assert_eq!(h.call("code", &[]).unwrap(), Value::Num(3.0));
        let sleeper = quick_shell(if cfg!(windows) {
            "ping -n 30 127.0.0.1 > nul"
        } else {
            "sleep 30"
        });
        assert_eq!(
            h.call("slow", &[Value::Str(sleeper)]).unwrap(),
            Value::Bool(true)
        );
    }

    #[test]
    fn the_sandbox_has_no_way_to_start_programs() {
        let h = LuaHost::sandboxed().unwrap();
        assert_eq!(
            h.eval("xl.run == nil and xl.shell == nil", &[]).unwrap(),
            Value::Bool(true)
        );
    }

    #[test]
    fn a_worker_call_is_stopped_by_its_deadline_and_by_cancellation() {
        let b = bundle_of(&[(
            "t.lua",
            "function spin() while true do end end function quick() return 7 end",
        )]);
        let h = LuaHost::new_worker(
            &b,
            &HostContext {
                workbook: Path::new("/tmp/book.xlsx"),
                cache_dir: None,
                xl: None,
            },
        )
        .unwrap();
        let cancel = Arc::new(AtomicBool::new(false));
        let soon = Limits {
            deadline: Instant::now() + Duration::from_millis(200),
            cancel: Arc::clone(&cancel),
        };
        let Err(HostError::Lua(m)) = h.call_limited("spin", &[], &soon) else {
            panic!("must time out")
        };
        assert!(m.contains("timed out"), "{m}");

        // The state is reusable afterwards, and a pre-set cancel flag aborts immediately.
        let fine = Limits {
            deadline: Instant::now() + Duration::from_secs(30),
            cancel: Arc::new(AtomicBool::new(false)),
        };
        assert_eq!(
            h.call_limited("quick", &[], &fine).unwrap(),
            Value::Num(7.0)
        );
        cancel.store(true, Ordering::Relaxed);
        let cancelled = Limits {
            deadline: Instant::now() + Duration::from_secs(30),
            cancel,
        };
        let Err(HostError::Lua(m)) = h.call_limited("spin", &[], &cancelled) else {
            panic!("must cancel")
        };
        assert!(m.contains("cancelled"), "{m}");
    }
}
