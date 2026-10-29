//! SQL over Excel data, powered by an in-memory SQLite database.
//!
//! A query names the data it needs and this module loads exactly that into temporary tables
//! before running the statement:
//!
//! * `[Sheet1!A1:C100]`, `[Orders]` (an Excel table, defined name or range): a bracketed reference
//!   in the SQL text. The first row of a range is the header. A reference that is a plain
//!   identifier becomes a table of that name (`select Orders.name from [Orders]`), anything
//!   else a generated `__rN`.
//! * Worksheet-function arguments: a range argument *k* becomes the table `tk`, a scalar argument
//!   *k* the parameter `?k`.
//!
//! The tables are snapshots (a query never changes the sheet), so everything here is plain data
//! and tested without Excel. Cell values map as: number -> REAL, text -> TEXT, boolean -> INTEGER
//! 0/1, empty -> NULL, error -> its text (`#N/A`). Dates stay numbers; `xl_iso()` / `xl_date()`
//! convert at the edges.

use std::path::PathBuf;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::{Duration, Instant};

use rusqlite::Connection;
use rusqlite::config::DbConfig;
use rusqlite::functions::FunctionFlags;
use rusqlite::types::{Value as SqlValue, ValueRef};

use crate::dates;
use crate::value::{Array, Value};
use crate::xlapi::{MAX_CELLS, Op, Reply, XlApi};

/// SQLite's default column limit is 2000.
const MAX_COLUMNS: usize = 2000;

/// Where bracketed references get their data.
pub trait Source {
    fn fetch(&self, reference: &str) -> Result<Array, String>;

    /// A database named by `[db:name]` in the SQL text, attached read-only for the query.
    fn database(&self, name: &str) -> Result<DbHandle, String> {
        Err(format!("[db:{name}]: databases are not available here"))
    }
}

/// What a `[db:name]` reference resolves to.
pub enum DbHandle {
    /// An embedded database (its bytes).
    Bytes(Vec<u8>),
    /// A database file on disk.
    File(PathBuf),
}

/// Looks databases up by name.
pub type DbResolver = Arc<dyn Fn(&str) -> Result<DbHandle, String> + Send + Sync>;

/// Adds database lookup to another source.
pub struct WithDatabases {
    pub base: Box<dyn Source>,
    pub databases: DbResolver,
}

impl Source for WithDatabases {
    fn fetch(&self, reference: &str) -> Result<Array, String> {
        self.base.fetch(reference)
    }
    fn database(&self, name: &str) -> Result<DbHandle, String> {
        (self.databases)(name)
    }
}

/// Reads references through Excel: an Excel table or defined name first, then a range address.
pub struct XlSource {
    pub api: Arc<dyn XlApi>,
    pub workbook: PathBuf,
}

impl Source for XlSource {
    fn fetch(&self, reference: &str) -> Result<Array, String> {
        if self.workbook.as_os_str().is_empty() {
            return NoSource.fetch(reference);
        }
        let as_array = |reply: Reply| match reply {
            Reply::Data(Value::Array(a)) => Some(a),
            Reply::Data(v) => Some(Array::new(1, 1, vec![v])),
            _ => None,
        };
        if !reference.contains('!') {
            match self.api.call(Op::Table {
                workbook: self.workbook.clone(),
                name: reference.to_owned(),
            }) {
                Ok(reply) => {
                    if let Some(a) = as_array(reply) {
                        return Ok(a);
                    }
                }
                Err(e) if e.retryable => return Err(format!("[{reference}]: {e}")),
                Err(_) => {} // not a table: try it as a range or name
            }
        }
        self.api
            .call(Op::Read {
                workbook: self.workbook.clone(),
                address: reference.to_owned(),
                formulas: false,
            })
            .map_err(|e| format!("[{reference}]: {e}"))
            .and_then(|r| as_array(r).ok_or_else(|| format!("[{reference}]: no data")))
    }
}

/// A source that has nothing (queries that only use their arguments).
pub struct NoSource;

impl Source for NoSource {
    fn fetch(&self, reference: &str) -> Result<Array, String> {
        Err(format!(
            "[{reference}]: references to cells need the workbook to be saved to a file"
        ))
    }
}

#[derive(Clone)]
pub struct Limits {
    pub cancel: Option<Arc<AtomicBool>>,
    pub timeout: Duration,
}

impl Default for Limits {
    fn default() -> Self {
        Limits {
            cancel: None,
            timeout: Duration::from_secs(60),
        }
    }
}

pub struct Query<'a> {
    pub sql: &'a str,
    /// Worksheet-style arguments: arrays become tables `t1`.., scalars the parameters `?1`..
    pub args: &'a [Value],
    /// Named parameters (`:name`, `@name`, `$name`).
    pub named: &'a [(String, Value)],
    pub source: &'a dyn Source,
    /// Refuse statements that change data.
    pub read_only: bool,
    pub limits: Limits,
}

#[derive(Debug, PartialEq)]
pub struct Output {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<Value>>,
    /// The statement could change data (it was not a read-only statement).
    pub modified: bool,
}

impl Output {
    /// Column names as the first row, then the data.
    pub fn with_header(self) -> Value {
        let cols = self.columns.len();
        let rows = self.rows.len() + 1;
        let mut data = Vec::with_capacity(rows * cols);
        data.extend(self.columns.into_iter().map(Value::Str));
        for r in self.rows {
            data.extend(r);
        }
        Value::Array(Array::new(rows, cols, data))
    }
}

// ------------------------------------------------------------------ rewriting the SQL text

/// Finds `[reference]` tokens outside strings, quoted identifiers and comments and replaces each
/// with a quoted table name. Returns the new text and (reference, table name) pairs.
fn rewrite(sql: &str, friendly_names: bool) -> (String, Vec<(String, String)>) {
    let chars: Vec<char> = sql.chars().collect();
    let mut out = String::with_capacity(sql.len());
    let mut refs: Vec<(String, String)> = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '\'' | '"' | '`' => {
                // string / quoted identifier; a doubled quote is an escaped quote
                out.push(c);
                i += 1;
                while i < chars.len() {
                    out.push(chars[i]);
                    if chars[i] == c {
                        if chars.get(i + 1) == Some(&c) {
                            out.push(c);
                            i += 2;
                            continue;
                        }
                        i += 1;
                        break;
                    }
                    i += 1;
                }
            }
            '-' if chars.get(i + 1) == Some(&'-') => {
                while i < chars.len() && chars[i] != '\n' {
                    out.push(chars[i]);
                    i += 1;
                }
            }
            '/' if chars.get(i + 1) == Some(&'*') => {
                out.push_str("/*");
                i += 2;
                while i < chars.len() && !(chars[i] == '*' && chars.get(i + 1) == Some(&'/')) {
                    out.push(chars[i]);
                    i += 1;
                }
                if i < chars.len() {
                    out.push_str("*/");
                    i += 2;
                }
            }
            '[' => match chars[i + 1..].iter().position(|&x| x == ']') {
                Some(len) => {
                    let token: String = chars[i + 1..i + 1 + len].iter().collect();
                    let token = token.trim().to_owned();
                    let name = match refs.iter().find(|(t, _)| *t == token) {
                        Some((_, n)) => n.clone(),
                        None => {
                            let n = if let Some(db) = token.strip_prefix("db:") {
                                format!("db_{}", db.trim())
                            } else if friendly_names && is_identifier(&token) {
                                token.clone()
                            } else {
                                format!("__r{}", refs.len() + 1)
                            };
                            refs.push((token, n.clone()));
                            n
                        }
                    };
                    out.push_str(&quote_ident(&name));
                    i += len + 2;
                }
                None => {
                    out.push(c);
                    i += 1;
                }
            },
            _ => {
                out.push(c);
                i += 1;
            }
        }
    }
    (out, refs)
}

/// Is there more than one statement? (A trailing `;` and comments after it do not count.)
fn has_several_statements(sql: &str) -> bool {
    let chars: Vec<char> = sql.chars().collect();
    let mut i = 0;
    let mut seen_semicolon = false;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '\'' | '"' | '`' | '[' => {
                if seen_semicolon {
                    return true;
                }
                let close = if c == '[' { ']' } else { c };
                i += 1;
                while i < chars.len() && chars[i] != close {
                    i += 1;
                }
            }
            '-' if chars.get(i + 1) == Some(&'-') => {
                while i < chars.len() && chars[i] != '\n' {
                    i += 1;
                }
            }
            '/' if chars.get(i + 1) == Some(&'*') => {
                i += 2;
                while i < chars.len() && !(chars[i] == '*' && chars.get(i + 1) == Some(&'/')) {
                    i += 1;
                }
                i += 1;
            }
            ';' => seen_semicolon = true,
            c if c.is_whitespace() => {}
            _ if seen_semicolon => return true,
            _ => {}
        }
        i += 1;
    }
    false
}

fn is_identifier(s: &str) -> bool {
    let mut it = s.chars();
    it.next()
        .is_some_and(|c| c.is_ascii_alphabetic() || c == '_')
        && it.all(|c| c.is_ascii_alphanumeric() || c == '_')
}

fn quote_ident(name: &str) -> String {
    format!("\"{}\"", name.replace('"', "\"\""))
}

// ------------------------------------------------------------------ loading data

fn to_sql(v: &Value) -> SqlValue {
    match v {
        Value::Empty => SqlValue::Null,
        Value::Num(n) => SqlValue::Real(*n),
        Value::Bool(b) => SqlValue::Integer(i64::from(*b)),
        Value::Str(s) => SqlValue::Text(s.clone()),
        Value::Err(e) => SqlValue::Text(e.excel_text().to_owned()),
        Value::Array(_) => SqlValue::Null,
    }
}

fn from_sql(v: ValueRef<'_>) -> Value {
    match v {
        ValueRef::Null => Value::Empty,
        ValueRef::Integer(i) => Value::Num(i as f64),
        ValueRef::Real(r) => Value::Num(r),
        ValueRef::Text(t) => Value::Str(String::from_utf8_lossy(t).into_owned()),
        ValueRef::Blob(b) => Value::Str(format!("<blob {} bytes>", b.len())),
    }
}

fn column_names(header: &[Value]) -> Vec<String> {
    let mut names: Vec<String> = Vec::with_capacity(header.len());
    for (i, h) in header.iter().enumerate() {
        let base = match h {
            Value::Str(s) if !s.trim().is_empty() => s.trim().to_owned(),
            Value::Num(n) => format!("{n}"),
            Value::Bool(b) => b.to_string(),
            Value::Err(e) => e.excel_text().to_owned(),
            _ => format!("col{}", i + 1),
        };
        let mut name = base.clone();
        let mut n = 2;
        while names.iter().any(|x| x.eq_ignore_ascii_case(&name)) {
            name = format!("{base}_{n}");
            n += 1;
        }
        names.push(name);
    }
    names
}

/// Creates the temp table `name` from `data` (first row = header) and fills it.
fn load_table(conn: &Connection, name: &str, data: &Array) -> Result<(), String> {
    if data.rows == 0 || data.cols == 0 {
        return Err(format!("{name}: no data"));
    }
    if data.cols > MAX_COLUMNS {
        return Err(format!("{name}: more than {MAX_COLUMNS} columns"));
    }
    let names = column_names(&data.data[..data.cols]);
    let cols = names
        .iter()
        .map(|n| quote_ident(n))
        .collect::<Vec<_>>()
        .join(", ");
    conn.execute_batch(&format!("CREATE TEMP TABLE {} ({cols})", quote_ident(name)))
        .map_err(|e| format!("{name}: {e}"))?;
    let marks = vec!["?"; data.cols].join(",");
    let tx = conn
        .unchecked_transaction()
        .map_err(|e| format!("{name}: {e}"))?;
    {
        let mut insert = tx
            .prepare(&format!(
                "INSERT INTO {} VALUES ({marks})",
                quote_ident(name)
            ))
            .map_err(|e| format!("{name}: {e}"))?;
        for row in data.data.chunks(data.cols).skip(1) {
            let vals: Vec<SqlValue> = row.iter().map(to_sql).collect();
            insert
                .execute(rusqlite::params_from_iter(vals.iter()))
                .map_err(|e| format!("{name}: {e}"))?;
        }
    }
    tx.commit().map_err(|e| format!("{name}: {e}"))
}

// ------------------------------------------------------------------ running a query

pub fn open() -> Result<Connection, String> {
    let conn = Connection::open_in_memory().map_err(|e| e.to_string())?;
    // Untrusted input hardening: no writable schema tricks, no triggers/views calling unsafe functions.
    conn.set_db_config(DbConfig::SQLITE_DBCONFIG_DEFENSIVE, true)
        .map_err(|e| e.to_string())?;
    conn.set_db_config(DbConfig::SQLITE_DBCONFIG_TRUSTED_SCHEMA, false)
        .map_err(|e| e.to_string())?;
    register_functions(&conn)?;
    Ok(conn)
}

pub fn register_functions(conn: &Connection) -> Result<(), String> {
    let flags = FunctionFlags::SQLITE_UTF8 | FunctionFlags::SQLITE_DETERMINISTIC;
    conn.create_scalar_function("xl_iso", 1, flags, |ctx| {
        Ok(match ctx.get_raw(0) {
            ValueRef::Integer(i) => dates::serial_to_iso(i as f64),
            ValueRef::Real(r) => dates::serial_to_iso(r),
            _ => None,
        })
    })
    .map_err(|e| e.to_string())?;
    conn.create_scalar_function("xl_date", 1, flags, |ctx| {
        Ok(match ctx.get_raw(0) {
            ValueRef::Text(t) => std::str::from_utf8(t).ok().and_then(dates::iso_to_serial),
            _ => None,
        })
    })
    .map_err(|e| e.to_string())
}

/// Runs `q` in a fresh in-memory database that is thrown away afterwards.
pub fn run(q: &Query<'_>) -> Result<Output, String> {
    let mut conn = open()?;
    run_on(&mut conn, q, true)
}

fn arm_limits(conn: &mut Connection, limits: &Limits) -> Result<(), String> {
    let deadline = Instant::now() + limits.timeout;
    let cancel = limits.cancel.clone();
    conn.progress_handler(
        1000,
        Some(move || {
            cancel.as_ref().is_some_and(|c| c.load(Ordering::Relaxed)) || Instant::now() > deadline
        }),
    )
    .map_err(|e| e.to_string())
}

/// Runs `q` on an existing connection. The cells and tables it refers to are loaded as temporary
/// tables for the duration of the call and dropped again. `friendly_names` lets a reference such
/// as `[Orders]` be called `Orders` in the SQL; persistent databases pass `false` so a loaded
/// table can never shadow one of their own.
pub fn run_on(
    conn: &mut Connection,
    q: &Query<'_>,
    friendly_names: bool,
) -> Result<Output, String> {
    arm_limits(conn, &q.limits)?;
    let result = run_inner(conn, q, friendly_names);
    let _ = conn.progress_handler(0, None::<fn() -> bool>);
    result
}

fn run_inner(conn: &mut Connection, q: &Query<'_>, friendly_names: bool) -> Result<Output, String> {
    let (sql, refs) = rewrite(q.sql, friendly_names);
    let mut created: Vec<String> = Vec::new();
    let mut attached: Vec<String> = Vec::new();
    let outcome = (|| {
        for (i, a) in q.args.iter().enumerate() {
            if let Value::Array(arr) = a {
                let name = format!("t{}", i + 1);
                load_table(conn, &name, arr)?;
                created.push(name);
            }
        }
        for (token, table) in &refs {
            if let Some(db) = token.strip_prefix("db:") {
                attach(conn, db.trim(), table, q.source.database(db.trim())?)?;
                attached.push(table.clone());
            } else {
                load_table(conn, table, &q.source.fetch(token)?)?;
                created.push(table.clone());
            }
        }
        execute(conn, &sql, q)
    })();
    // Leave a persistent connection as it was found.
    let _ = conn.pragma_update(None, "query_only", false);
    for name in &created {
        let _ = conn.execute_batch(&format!("DROP TABLE temp.{}", quote_ident(name)));
    }
    for name in &attached {
        let _ = conn.execute_batch(&format!("DETACH DATABASE {}", quote_ident(name)));
    }
    outcome
}

fn execute(conn: &Connection, sql: &str, q: &Query<'_>) -> Result<Output, String> {
    if q.read_only {
        conn.pragma_update(None, "query_only", true)
            .map_err(|e| e.to_string())?;
    }
    if has_several_statements(sql) {
        return Err("only one statement at a time here (xl.sql_exec / db:exec run several)".into());
    }
    let mut stmt = conn.prepare(sql).map_err(|e| e.to_string())?;
    let modified = !stmt.readonly();
    if q.read_only && modified {
        return Err(
            "this statement changes data, but SQL.QUERY only reads (use SQL.ASYNC or a macro)"
                .into(),
        );
    }
    let count = stmt.parameter_count();
    for (i, a) in q.args.iter().enumerate() {
        if i < count && !matches!(a, Value::Array(_)) {
            stmt.raw_bind_parameter(i + 1, to_sql(a))
                .map_err(|e| e.to_string())?;
        }
    }
    for (name, v) in q.named {
        let found = [":", "@", "$"]
            .iter()
            .find_map(|p| stmt.parameter_index(&format!("{p}{name}")).ok().flatten());
        match found {
            Some(idx) => stmt
                .raw_bind_parameter(idx, to_sql(v))
                .map_err(|e| e.to_string())?,
            None => return Err(format!("the statement has no parameter named '{name}'")),
        }
    }

    let columns: Vec<String> = stmt.column_names().map(str::to_owned).collect();
    if columns.is_empty() {
        stmt.raw_execute().map_err(interrupted)?;
        return Ok(Output {
            columns,
            rows: Vec::new(),
            modified,
        });
    }
    let mut rows = Vec::new();
    let mut cells: u64 = 0;
    let mut cursor = stmt.raw_query();
    while let Some(row) = cursor.next().map_err(interrupted)? {
        cells += columns.len() as u64;
        if cells > MAX_CELLS {
            return Err(format!("the result has more than {MAX_CELLS} cells"));
        }
        rows.push(
            (0..columns.len())
                .map(|i| from_sql(row.get_ref_unwrap(i)))
                .collect(),
        );
    }
    Ok(Output {
        columns,
        rows,
        modified,
    })
}

/// Attaches `handle` read-only as the schema `alias`.
fn attach(conn: &mut Connection, name: &str, alias: &str, handle: DbHandle) -> Result<(), String> {
    let a = quote_ident(alias);
    match handle {
        DbHandle::Bytes(bytes) => {
            conn.execute_batch(&format!("ATTACH DATABASE ':memory:' AS {a}"))
                .map_err(|e| format!("[db:{name}]: {e}"))?;
            conn.deserialize_read_exact(alias, bytes.as_slice(), bytes.len(), true)
                .map_err(|e| format!("[db:{name}]: {e}"))
        }
        DbHandle::File(path) => {
            let uri = file_uri(&path);
            conn.execute(&format!("ATTACH DATABASE ?1 AS {a}"), [uri])
                .map(drop)
                .map_err(|e| format!("[db:{name}]: {e}"))
        }
    }
}

/// `file:` URI for a path, opened read-only.
fn file_uri(path: &std::path::Path) -> String {
    let p = path.to_string_lossy().replace('\\', "/");
    let mut enc = String::new();
    for c in p.chars() {
        match c {
            '%' => enc.push_str("%25"),
            '?' => enc.push_str("%3f"),
            '#' => enc.push_str("%23"),
            ' ' => enc.push_str("%20"),
            _ => enc.push(c),
        }
    }
    let lead = if enc.starts_with('/') { "" } else { "/" };
    format!("file://{lead}{enc}?mode=ro")
}

fn interrupted(e: rusqlite::Error) -> String {
    match e {
        rusqlite::Error::SqliteFailure(f, _)
            if f.code == rusqlite::ErrorCode::OperationInterrupted =>
        {
            "the query was stopped (cancelled or timed out)".to_owned()
        }
        other => other.to_string(),
    }
}

/// Runs statements that return no rows (several separated by `;` are fine).
pub fn exec_script(sql: &str, source: &dyn Source, limits: &Limits) -> Result<(), String> {
    let mut conn = open()?;
    exec_script_on(&mut conn, sql, source, limits, true)
}

pub fn exec_script_on(
    conn: &mut Connection,
    sql: &str,
    source: &dyn Source,
    limits: &Limits,
    friendly_names: bool,
) -> Result<(), String> {
    arm_limits(conn, limits)?;
    let (sql, refs) = rewrite(sql, friendly_names);
    let mut created = Vec::new();
    let outcome = (|| {
        for (token, table) in &refs {
            load_table(conn, table, &source.fetch(token)?)?;
            created.push(table.clone());
        }
        conn.execute_batch(&sql).map_err(interrupted)
    })();
    for name in &created {
        let _ = conn.execute_batch(&format!("DROP TABLE temp.{}", quote_ident(name)));
    }
    let _ = conn.progress_handler(0, None::<fn() -> bool>);
    outcome
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    struct Fake(HashMap<&'static str, Array>);

    impl Source for Fake {
        fn fetch(&self, r: &str) -> Result<Array, String> {
            self.0
                .get(r)
                .cloned()
                .ok_or_else(|| format!("[{r}]: unknown"))
        }
    }

    fn s(x: &str) -> Value {
        Value::Str(x.into())
    }
    fn n(x: f64) -> Value {
        Value::Num(x)
    }

    fn orders() -> Array {
        Array::new(
            4,
            3,
            vec![
                s("name"),
                s("qty"),
                s("day"),
                s("bolt"),
                n(10.0),
                n(45_292.0),
                s("nut"),
                n(25.0),
                n(45_293.5),
                s("gear"),
                Value::Empty,
                n(45_351.0),
            ],
        )
    }

    fn source() -> Fake {
        let mut m = HashMap::new();
        m.insert("Orders", orders());
        m.insert(
            "Data!A1:B3",
            Array::new(
                3,
                2,
                vec![s("k"), s("v"), n(1.0), s("one"), n(2.0), s("two")],
            ),
        );
        Fake(m)
    }

    fn q<'a>(sql: &'a str, args: &'a [Value], src: &'a dyn Source) -> Query<'a> {
        Query {
            sql,
            args,
            named: &[],
            source: src,
            read_only: true,
            limits: Limits::default(),
        }
    }

    #[test]
    fn brackets_become_tables_and_headers_become_columns() {
        let src = source();
        let out = run(&q(
            "select name, qty from [Orders] where qty > 5 order by qty desc",
            &[],
            &src,
        ))
        .unwrap();
        assert_eq!(out.columns, ["name", "qty"]);
        assert_eq!(
            out.rows,
            vec![vec![s("nut"), n(25.0)], vec![s("bolt"), n(10.0)]]
        );
        // Qualified names work for table-like references; address ranges get generated names.
        let out = run(&q("select Orders.name, [Data!A1:B3].v from [Orders] join [Data!A1:B3] on [Data!A1:B3].k = 1", &[], &src)).unwrap();
        assert_eq!(out.rows.len(), 3);
        assert_eq!(out.rows[0][1], s("one"));
    }

    #[test]
    fn empty_cells_are_null_and_dates_convert() {
        let src = source();
        let out = run(&q(
            "select name, qty is null as blank, xl_iso(day) as d from [Orders] order by day",
            &[],
            &src,
        ))
        .unwrap();
        assert_eq!(out.rows[0], vec![s("bolt"), n(0.0), s("2024-01-01")]);
        assert_eq!(out.rows[2], vec![s("gear"), n(1.0), s("2024-02-29")]);
        let out = run(&q(
            "select xl_date('2024-02-29'), xl_date('nope'), xl_iso('x')",
            &[],
            &src,
        ))
        .unwrap();
        assert_eq!(out.rows[0], vec![n(45_351.0), Value::Empty, Value::Empty]);
    }

    #[test]
    fn arguments_are_tables_and_parameters() {
        let src = NoSource;
        let range = Value::Array(orders());
        let args = [range, n(20.0)];
        let out = run(&q("select name from t1 where qty > ?2", &args, &src)).unwrap();
        assert_eq!(out.rows, vec![vec![s("nut")]]);
        // Parameter numbers stay aligned with the argument positions even if only some are used.
        let out = run(&q("select ?2", &args, &src)).unwrap();
        assert_eq!(out.rows, vec![vec![n(20.0)]]);
    }

    #[test]
    fn several_statements_are_refused_not_silently_cut_off() {
        let src = source();
        let e = run(&q("select 1; select 2", &[], &src)).unwrap_err();
        assert!(e.contains("one statement"), "{e}");
        assert!(run(&q("select 1;", &[], &src)).is_ok());
        assert!(run(&q("select ';x' ; -- done\n", &[], &src)).is_ok());
    }

    #[test]
    fn named_parameters() {
        let src = source();
        let mut query = q("select name from [Orders] where qty >= :min", &[], &src);
        let named = [("min".to_owned(), n(25.0))];
        query.named = &named;
        assert_eq!(run(&query).unwrap().rows, vec![vec![s("nut")]]);
        let bad = [("nope".to_owned(), n(1.0))];
        query.named = &bad;
        assert!(run(&query).unwrap_err().contains("nope"));
    }

    #[test]
    fn brackets_in_strings_and_comments_are_left_alone() {
        let (sql, refs) = rewrite(
            "select '[not a ref]', \"[nor this]\" /* [nor] */ from [T] -- [x]\n",
            true,
        );
        assert_eq!(refs, vec![("T".to_owned(), "T".to_owned())]);
        assert!(sql.contains("'[not a ref]'") && sql.contains("-- [x]") && sql.contains("\"T\""));
    }

    #[test]
    fn read_only_refuses_changes_and_headers_are_cleaned() {
        let src = source();
        let e = run(&q("insert into [Orders] values ('x', 1, 2)", &[], &src)).unwrap_err();
        assert!(e.contains("only reads"), "{e}");
        let mut query = q("create table x(a)", &[], &src);
        query.read_only = false;
        assert!(run(&query).unwrap().rows.is_empty());
        let weird = Array::new(
            2,
            4,
            vec![
                s("a"),
                s("A"),
                Value::Empty,
                n(3.0),
                n(1.0),
                n(2.0),
                n(3.0),
                n(4.0),
            ],
        );
        let args = [Value::Array(weird)];
        let out = run(&q("select * from t1", &args, &NoSource)).unwrap();
        assert_eq!(out.columns, ["a", "A_2", "col3", "3"]);
    }

    #[test]
    fn errors_are_reported_and_runaway_queries_stop() {
        let src = source();
        assert!(
            run(&q("select nope from [Orders]", &[], &src))
                .unwrap_err()
                .contains("nope")
        );
        assert!(
            run(&q("select * from [Missing]", &[], &src))
                .unwrap_err()
                .contains("Missing")
        );
        let mut query = q(
            "with recursive r(x) as (select 1 union all select x + 1 from r) select count(*) from r",
            &[],
            &src,
        );
        query.limits.timeout = Duration::from_millis(200);
        let e = run(&query).unwrap_err();
        assert!(e.contains("stopped"), "{e}");
        let cancel = Arc::new(AtomicBool::new(true));
        query.limits = Limits {
            cancel: Some(cancel),
            timeout: Duration::from_secs(30),
        };
        assert!(run(&query).unwrap_err().contains("stopped"));
    }

    #[test]
    fn output_with_header_is_a_ready_to_write_table() {
        let src = source();
        let out = run(&q(
            "select name, qty from [Orders] order by name",
            &[],
            &src,
        ))
        .unwrap();
        let Value::Array(a) = out.with_header() else {
            panic!()
        };
        assert_eq!((a.rows, a.cols), (4, 2));
        assert_eq!(a.data[0], s("name"));
    }
}
