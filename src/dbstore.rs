//! Databases for Lua and SQL: embedded in the workbook, or files on the machine.
//!
//! A database name is looked up like a Lua module: first the workbook's embedded files, then the
//! workbook's own folder, then `%APPDATA%\xll-bridge\local` (which also holds machine Lua modules).
//! `sales` finds `sales.db`, `sales.sqlite` or `sales.sqlite3`; a full path always means that file.
//!
//! An **embedded** database lives in memory and is saved back into the workbook's file tree a
//! moment after it changes (the host collects those with [`take_flushes`]); saving the workbook
//! then stores it. Machine databases are ordinary SQLite files opened in place.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex, OnceLock};
use std::time::{Duration, Instant};

use rusqlite::{Connection, OpenFlags};
use sha2::{Digest, Sha256};

use crate::sqldb::{self, DbHandle, DbResolver, Limits, Output, Query, Source};
use crate::trust::normalize_path;
use crate::vfs::{self, Bundle};

/// How long a database must be quiet before it is written back into the workbook.
pub const FLUSH_AFTER: Duration = Duration::from_millis(1500);

const DB_EXTS: [&str; 3] = [".db", ".sqlite", ".sqlite3"];

/// `%APPDATA%\xll-bridge\local`: machine Lua modules and databases.
pub fn local_dir() -> Option<PathBuf> {
    crate::trust::TrustStore::default_path().and_then(|p| p.parent().map(|d| d.join("local")))
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Place {
    /// A file in the workbook's embedded tree.
    Embedded(String),
    /// A database file on disk.
    File(PathBuf),
}

fn has_db_extension(name: &str) -> bool {
    vfs::is_database_file(name)
}

fn candidates(name: &str) -> Vec<String> {
    if has_db_extension(name) {
        vec![name.to_owned()]
    } else {
        DB_EXTS
            .iter()
            .map(|e| format!("{name}{e}"))
            .chain(std::iter::once(name.to_owned()))
            .collect()
    }
}

fn is_absolute(name: &str) -> bool {
    Path::new(name).is_absolute()
        || name.starts_with("\\\\")
        || name.as_bytes().get(1) == Some(&b':')
}

fn relative_name(name: &str) -> Result<String, String> {
    vfs::clean_path(name).map_err(|e| format!("'{name}': {e}"))
}

/// Finds an existing database.
pub fn locate(
    name: &str,
    bundle: Option<&Bundle>,
    workbook: &Path,
    local: Option<&Path>,
) -> Result<Place, String> {
    let name = name.trim();
    if name.is_empty() {
        return Err("a database needs a name".into());
    }
    if is_absolute(name) {
        let p = PathBuf::from(name);
        return if p.is_file() {
            Ok(Place::File(p))
        } else {
            Err(format!("database file '{name}' does not exist"))
        };
    }
    let rel = relative_name(name)?;
    let cands = candidates(&rel);
    if let Some(b) = bundle {
        for c in &cands {
            if b.get(c).is_some() {
                return Ok(Place::Embedded(c.clone()));
            }
        }
    }
    // Created in this session and not stored in the workbook's files yet.
    for c in &cands {
        if registry().lock().unwrap().contains_key(&key(workbook, c)) {
            return Ok(Place::Embedded(c.clone()));
        }
    }
    let mut searched = vec!["the workbook's embedded files".to_owned()];
    for dir in [workbook.parent(), local].into_iter().flatten() {
        if dir.as_os_str().is_empty() {
            continue;
        }
        for c in &cands {
            let p = dir.join(c);
            if p.is_file() {
                return Ok(Place::File(p));
            }
        }
        searched.push(dir.display().to_string());
    }
    Err(format!(
        "database '{name}' not found (looked in {})",
        searched.join(", ")
    ))
}

/// Where a new database called `name` would be created: `embedded`, `local` or `folder` (next to
/// the workbook).
pub fn creation_place(
    name: &str,
    where_: &str,
    workbook: &Path,
    local: Option<&Path>,
) -> Result<Place, String> {
    let name = name.trim();
    if is_absolute(name) {
        return Ok(Place::File(PathBuf::from(name)));
    }
    let mut rel = relative_name(name)?;
    if !has_db_extension(&rel) {
        rel.push_str(".db");
    }
    match where_ {
        "embedded" => Ok(Place::Embedded(rel)),
        "local" => local
            .map(|d| Place::File(d.join(&rel)))
            .ok_or_else(|| "the machine folder is not available".into()),
        "folder" => workbook
            .parent()
            .filter(|d| !d.as_os_str().is_empty())
            .map(|d| Place::File(d.join(&rel)))
            .ok_or_else(|| "the workbook has no folder yet (save it first)".into()),
        other => Err(format!(
            "create must be \"embedded\", \"local\" or \"folder\", not \"{other}\""
        )),
    }
}

// ------------------------------------------------------------------ open databases

struct State {
    dirty_since: Option<Instant>,
    /// Digest of the database bytes as last loaded from / written to the workbook.
    synced: Option<String>,
}

pub struct Db {
    pub place: Place,
    workbook: PathBuf,
    conn: Mutex<Connection>,
    state: Mutex<State>,
    pub read_only: bool,
}

fn digest(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}

impl Db {
    /// Human-readable location.
    pub fn label(&self) -> String {
        match &self.place {
            Place::Embedded(p) => format!("embedded:{p}"),
            Place::File(p) => p.display().to_string(),
        }
    }

    fn is_embedded(&self) -> bool {
        matches!(self.place, Place::Embedded(_))
    }

    fn mark_dirty(&self) {
        if self.is_embedded() {
            self.state.lock().unwrap().dirty_since = Some(Instant::now());
        }
    }

    /// Runs one statement. `source` supplies `[range]` / `[db:name]` references.
    pub fn run(
        &self,
        sql: &str,
        args: &[crate::value::Value],
        named: &[(String, crate::value::Value)],
        source: &dyn Source,
        limits: Limits,
    ) -> Result<Output, String> {
        let mut conn = self.conn.lock().unwrap();
        let out = sqldb::run_on(
            &mut conn,
            &Query {
                sql,
                args,
                named,
                source,
                read_only: self.read_only,
                limits,
            },
            false,
        );
        if !matches!(&out, Ok(o) if !o.modified) {
            self.mark_dirty();
        }
        out
    }

    /// Runs statements that return no rows.
    pub fn exec(&self, script: &str, source: &dyn Source, limits: &Limits) -> Result<(), String> {
        if self.read_only {
            return Err("this database was opened read-only".into());
        }
        let mut conn = self.conn.lock().unwrap();
        let out = sqldb::exec_script_on(&mut conn, script, source, limits, false);
        self.mark_dirty();
        out
    }

    /// Replaces `table` with the contents of a CSV file (see [`crate::csvio`]).
    pub fn import_csv(
        &self,
        path: &Path,
        table: &str,
        opts: &crate::csvio::Options,
        limits: &Limits,
    ) -> Result<crate::csvio::Imported, String> {
        if self.read_only {
            return Err("this database was opened read-only".into());
        }
        let mut conn = self.conn.lock().unwrap();
        let out = crate::csvio::import(&mut conn, path, table, opts, limits);
        self.mark_dirty();
        out
    }

    /// The whole database as bytes (what gets stored in the workbook).
    pub fn snapshot(&self) -> Result<Vec<u8>, String> {
        let mut conn = self.conn.lock().unwrap();
        conn.serialize("main")
            .map(|d| d.to_vec())
            .map_err(|e| e.to_string())
    }
}

type Registry = Mutex<HashMap<(String, String), Arc<Db>>>;

fn registry() -> &'static Registry {
    static R: OnceLock<Registry> = OnceLock::new();
    R.get_or_init(Mutex::default)
}

fn key(workbook: &Path, path: &str) -> (String, String) {
    (normalize_path(workbook), path.to_ascii_lowercase())
}

fn memory_db(bytes: Option<&[u8]>) -> Result<Connection, String> {
    let mut conn = sqldb::open()?;
    if let Some(b) = bytes {
        conn.deserialize_read_exact("main", b, b.len(), false)
            .map_err(|e| format!("not a SQLite database: {e}"))?;
    }
    Ok(conn)
}

/// The embedded database `path` of `workbook`. `bytes` is what the workbook's file tree holds now
/// (`None` if there is no such file); with `create` a missing file starts as an empty database.
/// An already open database is shared by everything running in the workbook.
pub fn open_embedded(
    workbook: &Path,
    path: &str,
    bytes: Option<&[u8]>,
    create: bool,
    read_only: bool,
) -> Result<Arc<Db>, String> {
    let k = key(workbook, path);
    let mut reg = registry().lock().unwrap();
    if let Some(existing) = reg.get(&k) {
        let st = existing.state.lock().unwrap();
        let replaced = match (bytes, &st.synced) {
            (Some(b), Some(s)) => digest(b) != *s,
            _ => false,
        };
        // The file in the workbook was swapped for another one and nothing is waiting to be saved.
        if !(replaced && st.dirty_since.is_none()) {
            let db = Arc::clone(existing);
            drop(st);
            return Ok(if read_only && !db.read_only {
                Arc::new(read_only_view(&db)?)
            } else {
                db
            });
        }
    }
    let (conn, synced) = match bytes {
        Some(b) => (memory_db(Some(b))?, Some(digest(b))),
        None if create => (memory_db(None)?, None),
        None => return Err(format!("embedded database '{path}' does not exist")),
    };
    let db = Arc::new(Db {
        place: Place::Embedded(path.to_owned()),
        workbook: workbook.to_owned(),
        conn: Mutex::new(conn),
        state: Mutex::new(State {
            // A new database is stored in the workbook right away.
            dirty_since: bytes.is_none().then(Instant::now),
            synced,
        }),
        read_only: false,
    });
    reg.insert(k, Arc::clone(&db));
    drop(reg);
    if read_only {
        return Ok(Arc::new(read_only_view(&db)?));
    }
    Ok(db)
}

/// A read-only handle on the same data (a private snapshot).
fn read_only_view(db: &Db) -> Result<Db, String> {
    let bytes = db.snapshot()?;
    Ok(Db {
        place: db.place.clone(),
        workbook: db.workbook.clone(),
        conn: Mutex::new(memory_db(Some(&bytes))?),
        state: Mutex::new(State {
            dirty_since: None,
            synced: None,
        }),
        read_only: true,
    })
}

/// A database file on disk.
pub fn open_file(path: &Path, create: bool, read_only: bool) -> Result<Db, String> {
    if create {
        if let Some(dir) = path.parent() {
            std::fs::create_dir_all(dir).map_err(|e| format!("{}: {e}", dir.display()))?;
        }
    } else if !path.is_file() {
        return Err(format!("database file '{}' does not exist", path.display()));
    }
    let mut flags = OpenFlags::SQLITE_OPEN_NO_MUTEX;
    flags |= if read_only {
        OpenFlags::SQLITE_OPEN_READ_ONLY
    } else if create {
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE
    } else {
        OpenFlags::SQLITE_OPEN_READ_WRITE
    };
    let conn =
        Connection::open_with_flags(path, flags).map_err(|e| format!("{}: {e}", path.display()))?;
    use rusqlite::config::DbConfig;
    conn.set_db_config(DbConfig::SQLITE_DBCONFIG_DEFENSIVE, true)
        .and_then(|_| conn.set_db_config(DbConfig::SQLITE_DBCONFIG_TRUSTED_SCHEMA, false))
        .map_err(|e| e.to_string())?;
    conn.busy_timeout(Duration::from_secs(5))
        .map_err(|e| e.to_string())?;
    sqldb::register_functions(&conn)?;
    Ok(Db {
        place: Place::File(path.to_owned()),
        workbook: PathBuf::new(),
        conn: Mutex::new(conn),
        state: Mutex::new(State {
            dirty_since: None,
            synced: None,
        }),
        read_only,
    })
}

/// A change to store in the workbook.
pub struct Flush {
    pub workbook: PathBuf,
    /// Path inside the workbook's file tree.
    pub path: String,
    pub bytes: Vec<u8>,
    db: Arc<Db>,
}

impl Flush {
    /// Storing failed: try again later.
    pub fn failed(&self) {
        self.db.mark_dirty();
    }
}

/// Embedded databases that changed and have been quiet for `quiet`, serialised for storing.
pub fn take_flushes(quiet: Duration) -> Vec<Flush> {
    take_flushes_where(quiet, |_| true)
}

/// [`take_flushes`] for the databases of one workbook only.
pub fn take_flushes_for(workbook: &Path, quiet: Duration) -> Vec<Flush> {
    take_flushes_where(quiet, |w| w == workbook)
}

fn take_flushes_where(quiet: Duration, wanted: impl Fn(&Path) -> bool) -> Vec<Flush> {
    let dbs: Vec<Arc<Db>> = registry().lock().unwrap().values().cloned().collect();
    let mut out = Vec::new();
    for db in dbs {
        if !wanted(&db.workbook) {
            continue;
        }
        let due = db
            .state
            .lock()
            .unwrap()
            .dirty_since
            .is_some_and(|t| t.elapsed() >= quiet);
        if !due {
            continue;
        }
        let Place::Embedded(path) = &db.place else {
            continue;
        };
        match db.snapshot() {
            Ok(bytes) => {
                let mut st = db.state.lock().unwrap();
                st.dirty_since = None;
                st.synced = Some(digest(&bytes));
                out.push(Flush {
                    workbook: db.workbook.clone(),
                    path: path.clone(),
                    bytes,
                    db: Arc::clone(&db),
                });
            }
            Err(e) => crate::log::note(&format!("database {path}: cannot save: {e}")),
        }
    }
    out
}

/// Forgets the open embedded databases of a workbook (it was closed).
pub fn forget_workbook(workbook: &Path) {
    let w = normalize_path(workbook);
    registry().lock().unwrap().retain(|(wb, _), _| *wb != w);
}

// ------------------------------------------------------------------ `[db:name]` in SQL text

/// Resolves `[db:name]` references for a query: the current contents of an embedded database
/// (including changes not yet stored), or a machine file opened read-only.
pub fn resolver(workbook: PathBuf, bundle: Option<Arc<Bundle>>) -> DbResolver {
    Arc::new(move |name: &str| {
        let local = local_dir();
        match locate(name, bundle.as_deref(), &workbook, local.as_deref())? {
            Place::File(p) => Ok(DbHandle::File(p)),
            Place::Embedded(path) => {
                if let Some(db) = registry().lock().unwrap().get(&key(&workbook, &path)) {
                    return Ok(DbHandle::Bytes(db.snapshot()?));
                }
                bundle
                    .as_ref()
                    .and_then(|b| b.get(&path))
                    .map(|f| DbHandle::Bytes(f.data.clone()))
                    .ok_or_else(|| format!("embedded database '{path}' does not exist"))
            }
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::testutil::bundle_of;
    use crate::sqldb::NoSource;
    use crate::value::Value;

    fn run(db: &Db, sql: &str) -> Result<Output, String> {
        db.run(sql, &[], &[], &NoSource, Limits::default())
    }

    fn wb(n: &str) -> PathBuf {
        PathBuf::from(format!("/w/{n}.xlsx"))
    }

    #[test]
    fn names_are_found_like_lua_modules() {
        let dir = tempfile::tempdir().unwrap();
        let (book_dir, local) = (dir.path().join("book"), dir.path().join("local"));
        std::fs::create_dir_all(&book_dir).unwrap();
        std::fs::create_dir_all(&local).unwrap();
        let workbook = book_dir.join("b.xlsx");
        std::fs::write(book_dir.join("beside.db"), b"x").unwrap();
        std::fs::write(local.join("shared.sqlite"), b"x").unwrap();
        std::fs::write(local.join("beside.db"), b"x").unwrap();
        let mut bundle = bundle_of(&[("main.lua", "x = 1")]);
        bundle
            .put_file("data/sales.db", b"x".to_vec(), false)
            .unwrap();
        bundle.put_file("beside.db", b"x".to_vec(), false).unwrap();
        let find = |n: &str| locate(n, Some(&bundle), &workbook, Some(&local));
        assert_eq!(
            find("data/sales").unwrap(),
            Place::Embedded("data/sales.db".into())
        );
        // The embedded file wins over the same name on disk; then the workbook folder; then local.
        assert_eq!(find("beside").unwrap(), Place::Embedded("beside.db".into()));
        let no_embedded = |n: &str| locate(n, None, &workbook, Some(&local));
        assert_eq!(
            no_embedded("beside").unwrap(),
            Place::File(book_dir.join("beside.db"))
        );
        assert_eq!(
            find("shared").unwrap(),
            Place::File(local.join("shared.sqlite"))
        );
        let abs = local.join("shared.sqlite");
        assert_eq!(find(abs.to_str().unwrap()).unwrap(), Place::File(abs));
        let e = find("nope").unwrap_err();
        assert!(e.contains("not found") && e.contains("embedded"), "{e}");
        assert!(find("../x").is_err());
        assert_eq!(
            creation_place("log", "embedded", &workbook, Some(&local)).unwrap(),
            Place::Embedded("log.db".into())
        );
        assert_eq!(
            creation_place("log", "local", &workbook, Some(&local)).unwrap(),
            Place::File(local.join("log.db"))
        );
        assert!(creation_place("log", "cloud", &workbook, Some(&local)).is_err());
    }

    #[test]
    fn an_embedded_database_round_trips_through_the_workbook() {
        let w = wb("roundtrip");
        let db = open_embedded(&w, "t.db", None, true, false).unwrap();
        run(&db, "create table t(a, b)").unwrap();
        run(&db, "insert into t values (1, 'x'), (2, 'y')").unwrap();
        // Quiet long enough: serialised for the workbook, once.
        std::thread::sleep(Duration::from_millis(30));
        let flushes = take_flushes_for(&w, Duration::from_millis(10));
        let mine: Vec<_> = flushes.iter().collect();
        assert_eq!(mine.len(), 1);
        assert!(take_flushes_for(&w, Duration::from_millis(10)).is_empty());
        // Opening again from those bytes (a new session) finds the data.
        let bytes = mine[0].bytes.clone();
        forget_workbook(&w);
        let again = open_embedded(&w, "t.db", Some(&bytes), false, false).unwrap();
        let out = run(&again, "select a, b from t order by a").unwrap();
        assert_eq!(
            out.rows,
            vec![
                vec![Value::Num(1.0), Value::Str("x".into())],
                vec![Value::Num(2.0), Value::Str("y".into())]
            ]
        );
        assert!(!out.modified);
        // Reading does not make it dirty.
        std::thread::sleep(Duration::from_millis(30));
        assert!(take_flushes_for(&w, Duration::from_millis(10)).is_empty());
        forget_workbook(&w);
    }

    #[test]
    fn a_swapped_file_is_reloaded_unless_changes_are_waiting() {
        let w = wb("swap");
        let db = open_embedded(&w, "s.db", None, true, false).unwrap();
        db.exec(
            "create table a(x); insert into a values (1)",
            &NoSource,
            &Limits::default(),
        )
        .unwrap();
        std::thread::sleep(Duration::from_millis(30));
        let first = take_flushes_for(&w, Duration::from_millis(10))
            .pop()
            .unwrap();
        // Another file with the same name arrives (sync folder, panel): picked up.
        let other = {
            let c = memory_db(None).unwrap();
            c.execute_batch("create table b(y); insert into b values (7)")
                .unwrap();
            let mut c = c;
            c.serialize("main").unwrap().to_vec()
        };
        let swapped = open_embedded(&w, "s.db", Some(&other), false, false).unwrap();
        assert!(run(&swapped, "select y from b").is_ok());
        assert!(run(&swapped, "select x from a").is_err());
        // With unsaved changes pending the live database is kept.
        run(&swapped, "insert into b values (8)").unwrap();
        let kept = open_embedded(&w, "s.db", Some(&first.bytes), false, false).unwrap();
        assert_eq!(
            run(&kept, "select count(*) from b").unwrap().rows[0][0],
            Value::Num(2.0)
        );
        forget_workbook(&w);
    }

    #[test]
    fn read_only_views_cannot_change_anything() {
        let w = wb("ro");
        let db = open_embedded(&w, "r.db", None, true, false).unwrap();
        db.exec(
            "create table t(a); insert into t values (1)",
            &NoSource,
            &Limits::default(),
        )
        .unwrap();
        let ro = open_embedded(&w, "r.db", None, false, true).unwrap();
        assert!(run(&ro, "insert into t values (2)").is_err());
        assert_eq!(
            run(&ro, "select count(*) from t").unwrap().rows[0][0],
            Value::Num(1.0)
        );
        assert!(
            ro.exec("delete from t", &NoSource, &Limits::default())
                .is_err()
        );
        forget_workbook(&w);
    }

    #[test]
    fn machine_files_and_database_references_in_sql() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("m.db");
        let file = open_file(&path, true, false).unwrap();
        file.exec(
            "create table c(name, n); insert into c values ('a', 1), ('b', 2)",
            &NoSource,
            &Limits::default(),
        )
        .unwrap();
        drop(file);
        assert!(open_file(&dir.path().join("none.db"), false, false).is_err());

        // [db:name] attaches a machine file and an embedded database read-only to a query.
        let w = wb("attach");
        let embedded = open_embedded(&w, "e.db", None, true, false).unwrap();
        embedded
            .exec(
                "create table z(v); insert into z values (42)",
                &NoSource,
                &Limits::default(),
            )
            .unwrap();
        let mut bundle = Bundle::new();
        bundle
            .put_file("readme.txt", b"hi".to_vec(), false)
            .unwrap();
        let mut resolver_dirs = HashMap::new();
        resolver_dirs.insert("m", path.clone());
        let r = resolver(w.clone(), Some(Arc::new(bundle)));
        let source = sqldb::WithDatabases {
            base: Box::new(NoSource),
            databases: Arc::new({
                let r = Arc::clone(&r);
                let p = path.clone();
                move |n: &str| {
                    if n == "m" {
                        Ok(DbHandle::File(p.clone()))
                    } else {
                        r(n)
                    }
                }
            }),
        };
        let out = sqldb::run(&Query {
            sql: "select c.name, z.v from [db:m].c join [db:e].z where c.n = 2",
            args: &[],
            named: &[],
            source: &source,
            read_only: true,
            limits: Limits::default(),
        })
        .unwrap();
        assert_eq!(
            out.rows,
            vec![vec![Value::Str("b".into()), Value::Num(42.0)]]
        );
        let e = sqldb::run(&Query {
            sql: "select * from [db:missing].t",
            args: &[],
            named: &[],
            source: &source,
            read_only: true,
            limits: Limits::default(),
        })
        .unwrap_err();
        assert!(e.contains("missing"), "{e}");
        forget_workbook(&w);
    }

    #[test]
    fn loading_cells_into_a_database_does_not_shadow_its_tables() {
        use crate::value::Array;
        let w = wb("shadow");
        let db = open_embedded(&w, "c.db", None, true, false).unwrap();
        run(&db, "create table Customers(name)").unwrap();
        struct Cells;
        impl Source for Cells {
            fn fetch(&self, _: &str) -> Result<Array, String> {
                Ok(Array::new(
                    3,
                    1,
                    vec![
                        Value::Str("name".into()),
                        Value::Str("ann".into()),
                        Value::Str("bob".into()),
                    ],
                ))
            }
        }
        // The reference is a different table (temp, generated name), so this fills the real one.
        db.run(
            "insert into Customers select name from [Customers]",
            &[],
            &[],
            &Cells,
            Limits::default(),
        )
        .unwrap();
        let out = run(&db, "select count(*) from Customers").unwrap();
        assert_eq!(out.rows[0][0], Value::Num(2.0));
        // And nothing is left behind.
        let leftovers = run(&db, "select count(*) from sqlite_temp_master").unwrap();
        assert_eq!(leftovers.rows[0][0], Value::Num(0.0));
        forget_workbook(&w);
    }
}
