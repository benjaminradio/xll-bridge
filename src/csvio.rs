//! Reading CSV files (what `dscmd`, DAX Studio's command-line tool, and most exports produce)
//! into Excel values or straight into a SQLite table.
//!
//! Types are worked out from the first rows of each column: all-numeric columns become numbers
//! (REAL), the rest text, except that numbers with a leading zero (`007`) stay text. Columns named
//! in `dates` are read as ISO 8601 and stored as Excel date numbers, so dates sort and subtract
//! like everywhere else in this add-in. Empty fields are NULL / empty cells.

use std::fs::File;
use std::path::Path;
use std::time::Instant;

use rusqlite::Connection;
use rusqlite::types::Value as SqlValue;

use crate::dates;
use crate::sqldb::{self, Limits};
use crate::value::{Array, Value};

/// Rows looked at to decide each column's type.
const SAMPLE_ROWS: usize = 2000;

#[derive(Clone, Debug, Default)]
pub struct Options {
    /// Turn DAX-style headers (`'Dim BU'[Business Unit]`, `[Total Sales]`) into plain names.
    pub clean_names: bool,
    /// Columns (by name, after cleaning) holding ISO 8601 dates.
    pub dates: Vec<String>,
    /// Columns that must stay text even if they look numeric (codes, account numbers).
    pub text: Vec<String>,
    /// Columns to index after loading.
    pub index: Vec<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Num,
    Text,
    Date,
}

#[derive(Debug, PartialEq)]
pub struct Imported {
    pub rows: u64,
    pub columns: Vec<String>,
}

/// `'Table'[Column]` and `Table[Column]` become `Column`; `[Measure]` becomes `Measure`. Where
/// that would make two columns share a name, those keep their table: `Table_Column`.
pub fn clean_header(raw: &[String], clean: bool) -> Vec<String> {
    let strip_bom = |s: &str| s.trim_start_matches('\u{feff}').trim().to_owned();
    if !clean {
        return raw.iter().map(|s| strip_bom(s)).collect();
    }
    let parts: Vec<(String, String)> = raw
        .iter()
        .map(|s| {
            let s = strip_bom(s);
            match (s.rfind('['), s.ends_with(']')) {
                (Some(open), true) => {
                    let col = s[open + 1..s.len() - 1].to_owned();
                    let table = s[..open].trim().trim_matches('\'').to_owned();
                    (table, col)
                }
                _ => (String::new(), s),
            }
        })
        .collect();
    parts
        .iter()
        .map(|(table, col)| {
            let dupes = parts
                .iter()
                .filter(|(_, c)| c.eq_ignore_ascii_case(col))
                .count();
            if dupes > 1 && !table.is_empty() {
                format!("{table}_{col}")
            } else {
                col.clone()
            }
        })
        .collect()
}

fn looks_numeric(field: &str) -> bool {
    let f = field.trim();
    if f.is_empty() {
        return true; // blanks do not decide
    }
    // "007" and "0123456" are codes, not numbers; "0", "0.5" and "-0.5" are numbers.
    let digits = f.trim_start_matches(['-', '+']);
    if digits.len() > 1 && digits.starts_with('0') && !digits.starts_with("0.") {
        return false;
    }
    f.parse::<f64>().is_ok_and(f64::is_finite)
}

fn convert(kind: Kind, field: &str) -> Value {
    if field.is_empty() {
        return Value::Empty;
    }
    match kind {
        Kind::Text => Value::Str(field.to_owned()),
        Kind::Num => field
            .trim()
            .parse::<f64>()
            .map_or_else(|_| Value::Str(field.to_owned()), Value::Num),
        Kind::Date => {
            dates::iso_to_serial(field).map_or_else(|| Value::Str(field.to_owned()), Value::Num)
        }
    }
}

fn reader(path: &Path) -> Result<csv::Reader<File>, String> {
    csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_path(path)
        .map_err(|e| format!("{}: {e}", path.display()))
}

struct Layout {
    columns: Vec<String>,
    kinds: Vec<Kind>,
}

/// Header and column types, from the first rows.
fn layout(path: &Path, opts: &Options) -> Result<Layout, String> {
    let mut rd = reader(path)?;
    let mut records = rd.records();
    let header = records
        .next()
        .ok_or_else(|| format!("{}: the file is empty", path.display()))?
        .map_err(|e| e.to_string())?;
    let raw: Vec<String> = header.iter().map(str::to_owned).collect();
    let columns = clean_header(&raw, opts.clean_names);
    let named = |list: &[String], c: &str| list.iter().any(|d| d.eq_ignore_ascii_case(c));
    let mut kinds: Vec<Kind> = columns
        .iter()
        .map(|c| {
            if named(&opts.dates, c) {
                Kind::Date
            } else if named(&opts.text, c) {
                Kind::Text
            } else {
                Kind::Num
            }
        })
        .collect();
    for rec in records.take(SAMPLE_ROWS) {
        let rec = rec.map_err(|e| e.to_string())?;
        for (i, field) in rec.iter().enumerate().take(kinds.len()) {
            if kinds[i] == Kind::Num && !looks_numeric(field) {
                kinds[i] = Kind::Text;
            }
        }
    }
    Ok(Layout { columns, kinds })
}

/// The whole file as a table of values, header row first. Refuses files larger than `max_cells`.
pub fn read_array(path: &Path, opts: &Options, max_cells: u64) -> Result<Array, String> {
    let lay = layout(path, opts)?;
    let cols = lay.columns.len();
    let mut data: Vec<Value> = lay.columns.iter().cloned().map(Value::Str).collect();
    let mut rd = reader(path)?;
    for rec in rd.records().skip(1) {
        let rec = rec.map_err(|e| e.to_string())?;
        for i in 0..cols {
            data.push(convert(lay.kinds[i], rec.get(i).unwrap_or("")));
        }
        if data.len() as u64 > max_cells {
            return Err(format!(
                "{} has more than {max_cells} cells; load it into a database instead",
                path.display()
            ));
        }
    }
    Ok(Array::new(data.len() / cols.max(1), cols, data))
}

fn to_sql(v: &Value) -> SqlValue {
    sqldb::to_sql(v)
}

/// Replaces the table `table` with the file's contents, in one transaction: other readers keep
/// seeing the old data until the new data is complete, and a failure leaves the old table alone.
pub fn import(
    conn: &mut Connection,
    path: &Path,
    table: &str,
    opts: &Options,
    limits: &Limits,
) -> Result<Imported, String> {
    let lay = layout(path, opts)?;
    let cols = lay.columns.len();
    if cols == 0 {
        return Err(format!("{}: no columns", path.display()));
    }
    let names = sqldb::column_names(
        &lay.columns
            .iter()
            .cloned()
            .map(Value::Str)
            .collect::<Vec<_>>(),
    );
    let q = sqldb::quote_ident(table);
    let col_list = names
        .iter()
        .map(|n| sqldb::quote_ident(n))
        .collect::<Vec<_>>()
        .join(", ");
    let marks = vec!["?"; cols].join(",");
    let deadline = Instant::now() + limits.timeout;

    let tx = conn.transaction().map_err(|e| e.to_string())?;
    tx.execute_batch(&format!(
        "DROP TABLE IF EXISTS {q}; CREATE TABLE {q} ({col_list});"
    ))
    .map_err(|e| format!("{table}: {e}"))?;
    let mut rows = 0u64;
    {
        let mut insert = tx
            .prepare(&format!("INSERT INTO {q} VALUES ({marks})"))
            .map_err(|e| e.to_string())?;
        let mut rd = reader(path)?;
        for rec in rd.records().skip(1) {
            let rec = rec.map_err(|e| format!("row {}: {e}", rows + 2))?;
            let vals: Vec<SqlValue> = (0..cols)
                .map(|i| to_sql(&convert(lay.kinds[i], rec.get(i).unwrap_or(""))))
                .collect();
            insert
                .execute(rusqlite::params_from_iter(vals.iter()))
                .map_err(|e| e.to_string())?;
            rows += 1;
            if rows.is_multiple_of(20_000)
                && (limits
                    .cancel
                    .as_ref()
                    .is_some_and(|c| c.load(std::sync::atomic::Ordering::Relaxed))
                    || Instant::now() > deadline)
            {
                return Err(
                    "the import was stopped (cancelled or timed out); the old table is unchanged"
                        .into(),
                );
            }
        }
    }
    for col in &opts.index {
        let Some(name) = names.iter().find(|n| n.eq_ignore_ascii_case(col)) else {
            return Err(format!("cannot index '{col}': no such column"));
        };
        tx.execute_batch(&format!(
            "CREATE INDEX {} ON {q} ({})",
            sqldb::quote_ident(&format!("{table}_{name}_idx")),
            sqldb::quote_ident(name)
        ))
        .map_err(|e| e.to_string())?;
    }
    tx.commit().map_err(|e| e.to_string())?;
    Ok(Imported {
        rows,
        columns: names,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn file(dir: &Path, name: &str, text: &str) -> std::path::PathBuf {
        let p = dir.join(name);
        std::fs::write(&p, text).unwrap();
        p
    }

    #[test]
    fn dax_headers_become_plain_names() {
        let h = |v: &[&str], clean| {
            clean_header(&v.iter().map(|s| s.to_string()).collect::<Vec<_>>(), clean)
        };
        assert_eq!(
            h(
                &[
                    "'Dim BU'[Business Unit]",
                    "Accounts[Code]",
                    "[Total Sales]",
                    "Plain"
                ],
                true
            ),
            ["Business Unit", "Code", "Total Sales", "Plain"]
        );
        // Same column name from two tables: keep the table to tell them apart.
        assert_eq!(
            h(&["BU[Name]", "Account[Name]", "[Amount]"], true),
            ["BU_Name", "Account_Name", "Amount"]
        );
        assert_eq!(h(&["\u{feff}a", " b "], false), ["a", "b"]);
    }

    #[test]
    fn types_come_from_the_data() {
        let d = tempfile::tempdir().unwrap();
        let p = file(
            d.path(),
            "t.csv",
            "\u{feff}Month,Account,Amount,Note\n2024-01-31,007,10.5,\"a, b\"\n2024-02-29,008,-3,\n",
        );
        let opts = Options {
            dates: vec!["Month".into()],
            ..Default::default()
        };
        let a = read_array(&p, &opts, 1000).unwrap();
        assert_eq!((a.rows, a.cols), (3, 4));
        assert_eq!(a.data[0], Value::Str("Month".into()));
        assert_eq!(a.data[4], Value::Num(45_322.0)); // 2024-01-31
        assert_eq!(a.data[5], Value::Str("007".into())); // leading zero: a code
        assert_eq!(a.data[6], Value::Num(10.5));
        assert_eq!(a.data[7], Value::Str("a, b".into()));
        assert_eq!(a.data[11], Value::Empty);
        assert!(
            read_array(&p, &opts, 5)
                .unwrap_err()
                .contains("more than 5 cells")
        );
        assert!(read_array(&d.path().join("none.csv"), &opts, 5).is_err());
    }

    #[test]
    fn importing_replaces_the_table_atomically_and_indexes_it() {
        let d = tempfile::tempdir().unwrap();
        let mut conn = Connection::open_in_memory().unwrap();
        let lim = Limits::default();
        let one = file(
            d.path(),
            "one.csv",
            "'Dim'[BU],[Sales]\nNorth,1\nSouth,2\nNorth,3\n",
        );
        let opts = Options {
            clean_names: true,
            index: vec!["BU".into()],
            ..Default::default()
        };
        let r = import(&mut conn, &one, "sales", &opts, &lim).unwrap();
        assert_eq!(
            (r.rows, r.columns.clone()),
            (3, vec!["BU".into(), "Sales".into()])
        );
        let total: f64 = conn
            .query_row("select sum(Sales) from sales where BU = 'North'", [], |r| {
                r.get(0)
            })
            .unwrap();
        assert_eq!(total, 4.0);
        let indexes: i64 = conn
            .query_row(
                "select count(*) from sqlite_master where type = 'index'",
                [],
                |r| r.get(0),
            )
            .unwrap();
        assert_eq!(indexes, 1);

        // A broken second file leaves the first import intact.
        let bad = file(d.path(), "bad.csv", "BU,Sales\nNorth,1\n");
        let bad_opts = Options {
            index: vec!["Nope".into()],
            ..Default::default()
        };
        assert!(import(&mut conn, &bad, "sales", &bad_opts, &lim).is_err());
        let rows: i64 = conn
            .query_row("select count(*) from sales", [], |r| r.get(0))
            .unwrap();
        assert_eq!(rows, 3);
        // A good one replaces it.
        let two = file(d.path(), "two.csv", "BU,Sales\nEast,9\n");
        import(&mut conn, &two, "sales", &Options::default(), &lim).unwrap();
        let rows: i64 = conn
            .query_row("select count(*) from sales", [], |r| r.get(0))
            .unwrap();
        assert_eq!(rows, 1);
    }

    #[test]
    fn a_few_million_rows_load_in_reasonable_time() {
        let d = tempfile::tempdir().unwrap();
        let p = d.path().join("big.csv");
        {
            use std::io::Write;
            let mut f = std::io::BufWriter::new(File::create(&p).unwrap());
            writeln!(f, "Month,BU,Account,Book,Amount").unwrap();
            for i in 0..300_000u32 {
                writeln!(
                    f,
                    "2024-{:02}-28,BU{},{},B{},{}.5",
                    i % 12 + 1,
                    i % 45,
                    i % 400,
                    i % 5,
                    i
                )
                .unwrap();
            }
        }
        let mut conn = Connection::open_in_memory().unwrap();
        let start = Instant::now();
        let opts = Options {
            dates: vec!["Month".into()],
            index: vec!["BU".into()],
            ..Default::default()
        };
        let r = import(&mut conn, &p, "gl", &opts, &Limits::default()).unwrap();
        assert_eq!(r.rows, 300_000);
        // Generous: debug builds are slow; 3.5 million rows is about twelve times this.
        assert!(start.elapsed().as_secs() < 60, "{:?}", start.elapsed());
    }
}
