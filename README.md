# xll-bridge

A native Excel add-in (XLL, written in Rust) that puts Lua 5.4 ([mlua]) behind worksheet functions.
Your Lua scripts, modules, data files and even DLLs live **inside the workbook** and stay **disabled
until you enable them**, like VBA macros. Lua can read and write cells, run external programs, work
in the background and stream live values into cells.

* **Windows x64, Excel 2010+** (the add-in is an XLL plus a small COM add-in in the same DLL).
* Built for an internal tool where **you write the scripts**: once a workbook is enabled its scripts
  have the full Lua standard library (see [Trust model](#trust-model)).
* The portable core is unit-tested on Linux; the Excel layer is tested in Excel by hand and its
  Windows-only parts run in CI on Windows. CI builds the downloadable package (`.xll`, `xllb-pack.exe`,
  `demo.xlsx`, `demo.lua`, checksums).

## Contents

[Quick start](#quick-start) · [Worksheet functions](#worksheet-functions) · [Lua API](#lua-api) · [SQL over cells](#sql-over-cells) ·
[The Lua ribbon tab](#the-lua-ribbon-tab) · [Editing scripts in your own editor](#editing-scripts-in-your-own-editor)
· [Trust model](#trust-model) · [Files on disk](#files-on-disk) · [How it works](#how-it-works) ·
[Troubleshooting](#troubleshooting) · [Limitations](#limitations) · [Future considerations](#future-considerations)

## Quick start

1. Download the package from the latest CI run (or a release), unblock the files
   (right-click ▸ Properties ▸ Unblock) and open `xll-bridge.xll` in Excel (File ▸ Open, or
   File ▸ Options ▸ Add-ins ▸ Go ▸ Browse). A **Lua** tab appears on the ribbon.
2. Open `demo.xlsx`. Its scripts are **blocked**: the Lua tab's status button is red. Click it,
   read the list of scripts and answer **Yes** to enable them.
3. Try the formulas:

```
=LUA.CALL("add", A1, B1)         calls the global Lua function add(a, b)
=LUA.EVAL("... * 2", A1)         evaluates an expression or chunk; arguments arrive as ...
=LUA.ASYNC("slow", A1)           like LUA.CALL, but in the background
=LUA.STREAM("clock")             a cell that keeps updating
=LUA.STATUS()                    scripts: none / blocked / enabled
```

To put your own scripts into a workbook, either use the **Files** panel / **Sync folder** on the Lua tab,
or the command line (works without Excel):

```sh
cargo build --release                       # copy target/release/xll_bridge.dll to xll-bridge.xll
cargo run --bin xllb-pack -- embed Book.xlsx Book-lua.xlsx main.lua lib/   # files and folders
cargo run --bin xllb-pack -- list Book-lua.xlsx
cargo run --bin xllb-pack -- extract Book-lua.xlsx ./out
```

Workbooks must be saved to a local path (OneDrive/SharePoint copies are resolved through Excel's
`LocalFullName`).

## Worksheet functions

| Function | Runs on | Cells | Use for |
|---|---|---|---|
| `LUA.CALL(name, a1…a10)` | Excel's UI thread | read only | quick calculations |
| `LUA.EVAL(code, a1…a10)` | Excel's UI thread | read only | one-off expressions (sandboxed when the workbook has no enabled scripts) |
| `LUA.ASYNC(name, a1…a10)` | background worker | read and write | slow work: web requests, programs |
| `LUA.STREAM(name, a1…a10)` | its own thread | read and write | values that change over time |
| `SQL.QUERY(sql, a1…a10)` | Excel's UI thread | read only | SQL over ranges and Excel tables ([below](#sql-over-cells)) |
| `SQL.ASYNC(sql, a1…a10)` | background worker | read only | the same for slow queries |
| `LUA.STATUS()` | | | trust state of this workbook (and, once streams are used, RTD diagnostics) |

`LUA.CALL` runs on Excel's UI thread, so a slow function freezes Excel while it runs; use `LUA.ASYNC`
for anything that takes more than a moment.

### `LUA.ASYNC`: slow calls

The cell shows `#GETTING_DATA` (what Excel's own async functions show) and fills in when the result
arrives; only the waiting cells recalculate and Excel stays usable.

* Each worker has its **own Lua state**, built from the workbook's enabled files: globals set by one call
  are not visible to `LUA.CALL` or to other workers. Pass what a function needs as arguments.
* Results are **cached** per function, arguments and script version, so recalculating does not re-run a
  slow call. They run again when the arguments or scripts change, or with **Re-run async** on the Lua tab.
* A call is limited to 120 s and is cancelled when re-run or when the workbook closes (a Lua loop is
  interrupted; a program started with `xl.run` is killed).
* In manual calculation mode finished cells are marked dirty and appear on the next F9.

### `LUA.STREAM`: live data

```lua
function clock()                 -- =LUA.STREAM("clock")
  while true do
    xl.emit(os.date("%H:%M:%S")) -- push a value; the cell updates
    xl.sleep(1)                  -- wait without blocking Excel (ends promptly when cancelled)
  end
end
```

The function runs on its own thread for as long as some cell shows it. `xl.emit(value)` pushes a single
value (not a table); a value returned at the end is pushed too, and an error ends the stream and shows
`#LUA: …`. Cells with the same function and arguments share one stream, and only the latest value
counts, so a fast stream cannot build a backlog. It stops when no cell uses it, when the workbook's
scripts are disabled, or when Excel exits.

Streams use Excel's RTD mechanism, and Excel batches real-time updates to its *throttle interval*
(2 seconds by default, so a stream that emits every second shows a new value every 2 seconds). Lower it
with `Application.RTD.ThrottleInterval = 500` in the VBA immediate window (a per-user Excel setting).

## Lua API

Conventions: a range is a table of rows (`{{1,2},{3,4}}`); a flat list returns a column; an empty cell is
`nil` (`xl.EMPTY` inside tables); errors are `xl.NA`, `xl.VALUE`, `xl.DIV0`, …

### Cells

```lua
local r = xl.range("Sheet1!A1:C10")   -- resolved once; "A1:C10" alone means the active sheet
r:value()          -- a bare value for one cell, otherwise a table of rows
r:values()         -- always a table of rows
r:formulas()       -- formulas (constants come back as their values)
r:set(42)          -- a scalar fills the whole range
r:set({{1,2},{3,4}})   -- a table is written from the top-left cell, at its own size
r:set_formulas({{"=A1+1"}})
r:clear()          -- ClearContents
r:size()           -- rows, cols          r:address()   -- "Sheet1!$A$1:$C$10"
xl.sheets()  xl.used_range("Sheet1")  xl.selection()
```

Addresses may be `Sheet!A1`, `'My Sheet'!A1`, `[Other.xlsx]Sheet!A1` or a defined name. Multi-area ranges
are not supported; one operation is limited to 5,000,000 cells.

* **`LUA.CALL` and `LUA.EVAL` can read but not write**: Excel does not let a worksheet function change
  cells. Reads there are not dependency-tracked, so pass a range as an argument if the result must
  recalculate when it changes.
* **`LUA.ASYNC`, `LUA.STREAM` and macros can read and write.** Their code runs on a worker thread and
  each range operation is handed to Excel's UI thread, so batch reads and writes (one `values()` / `set()`
  of a block, not one call per cell). If Excel is busy (editing a cell, a dialog) an operation is retried
  for a while, then fails with an error.

### Macros

`xl.macro("Name", function() … end)` registers a function for the ribbon's **Lua ▸ Macros** menu (see
below). A macro runs in the background in a fresh Lua state, may change cells, and is stopped after
5 minutes.

### Running programs

```lua
local r = xl.run{ "git", "status", "--short", cwd = "C:\\work\\repo", timeout = 10 }
-- r = { code, stdout, stderr, ok, timed_out, truncated } ; or nil, "message" if it could not start
local s = xl.shell("dir /b | find \".lua\"", { timeout = 5 })   -- through cmd /C (sh -c elsewhere)
```

Output is captured (32 MB cap per stream), `stdin = "text"` is supported, and `timeout` (seconds, default 60)
really kills the process (and its children on Windows). No console window appears. Cancelling a
background call also kills the program it is waiting for. These functions exist only in enabled workbooks.

### Files, modules and DLLs

* `.lua` files in the **root** of the embedded tree run (in name order) when the workbook is enabled and
  define the globals that `LUA.CALL` can call.
* Files in **folders** are modules: `require("lib.util")` loads `lib/util.lua` (or `lib/util/init.lua`).
* `xl.read_file("data/x.csv")` returns any embedded file's bytes as a string; `xl.files()` lists them.
* `.dll` / `.so` files are extracted (only once the workbook is enabled) to
  `%LOCALAPPDATA%\xll-bridge\cache\<digest>\` and added to `package.cpath`, so `require("native.thing")`
  finds `native/thing.dll`.
* `require` also searches the workbook's own folder and the machine folder `%APPDATA%\xll-bridge\local` (shared Lua and native modules, and machine databases).
* Paths are validated everywhere (no `..`, drive letters, reserved Windows names); text files are stored
  with `\n` line endings; limits are 32 MB per file and 64 MB per workbook.

## SQL over cells

An embedded SQLite engine (no database server, nothing to install) runs SQL over your cells. SQL needs no
enabled scripts, so `SQL.QUERY` and `SQL.ASYNC` work in any workbook.

```
=SQL.QUERY("select item, qty from t1 where qty > ?2 order by qty desc", A16:B19, 8)
=SQL.QUERY("select c.name, sum(o.qty) from [Orders] o join [Customers] c on c.id = o.cust group by 1")
=SQL.QUERY("select xl_iso(Day), Amount from [Sales!A1:B400] where Day >= xl_date('2024-01-01')")
```

* **Arguments**: a range argument *k* becomes the table `tk` (its first row is the header row and gives the
  column names); a scalar argument *k* is the parameter `?k`. Range arguments are tracked by Excel, so the
  formula recalculates when they change.
* **References in the SQL text**: `[Sheet1!A1:C9]` (first row = header), `[Orders]` (an Excel table or
  defined name). A reference that is a plain name becomes a table of that name, so `[Orders]` can be used as
  `Orders.qty`; anything else gets a generated name, so give it an alias (`[Sheet1!A1:C9] as a`). These reads
  are not dependency-tracked: pass a range as an argument when the result must refresh by itself.
* **Result**: the column names, then the rows, spilling from the formula cell (Excel 365). `SQL.QUERY`
  accepts one statement that only reads; `SQL.ASYNC` is the same on a worker (`#GETTING_DATA` until done).
* **Types**: numbers, text and empty cells (NULL) map as you would expect; booleans are 0/1; errors become
  their text (`#N/A`). Tables are snapshots: a query never changes the sheet.
* **Dates stay numbers** (they sort and subtract correctly). Convert at the edges: SQL `xl_iso(n)` /
  `xl_date('2024-02-29')`, Lua `xl.date_to_iso(n)` / `xl.iso_to_date(s)`. ISO 8601 text is
  `YYYY-MM-DD`, with `THH:MM:SS[.mmm]` when there is a time; the 1900 date system is assumed.
* Hardening: the schema is not trusted, untrusted-schema tricks are off, extensions cannot be loaded and a
  query is stopped after 60 seconds (or when its background call is cancelled).

### Databases

SQLite is also a place to keep data. A database can live **in the workbook** (travels with the file) or
**on the machine** (persistent, shared by every workbook). Names are found like Lua modules: first the
workbook's embedded files, then the workbook's own folder, then `%APPDATA%\xll-bridge\local`. `sales`
finds `sales.db`, `sales.sqlite` or `sales.sqlite3`; a full path means exactly that file.

```lua
local db = xl.db.open("sales")                          -- embedded, beside the workbook or machine, in that order
local db = xl.db.open("log", { create = "embedded" })   -- or "local" (machine folder) or "folder" (beside the workbook)
db:exec("create table if not exists orders(id integer primary key, item, qty)")
db:sql("insert into orders(item, qty) values (?, ?)", { "bolt", 10 })
local rows = db:sql("select item, qty from orders")      -- table of rows, column names first
local recs = db:records("select * from orders where qty > :n", { n = 5 })
db:sql("insert into orders(item, qty) select item, qty from [Sheet1!A1:B20]")   -- cells into the database
xl.db.list()                                            -- { {name=..., place="embedded"|"folder"|"local"}, ... }
```

* `db:sql` runs one statement, `db:exec` any number that return no rows. `[Sheet1!A1:B20]` and `[Table]`
  load cells for that one statement (as a temporary table with a generated name, so they never shadow the
  database's own tables). `{ readonly = true }` opens a database that cannot be changed.
* An **embedded** database is kept in memory and stored back into the workbook's file tree about a second and
  a half after the last change (the status bar says so); the workbook is marked modified, so **save it to
  keep the data**. Everything running in the workbook (`LUA.CALL`, background calls, macros, streams) shares
  one live copy. The database must fit the 32 MB file limit; use a machine database for big data.
* Embedded databases are only available in workbooks whose scripts are **enabled**. A database is data, not
  code: swapping one does not change what a workbook is trusted to run, but SQLite is opened defensively
  (no unsafe schema functions, no extensions).
* A database file in a folder you **Sync** is part of that folder; edits made by macros live in the workbook
  copy until the folder's file itself changes.
* In `SQL.QUERY` / `SQL.ASYNC` (read only) use `[db:name]` as a schema: `select * from [db:sales].orders`.
  Embedded databases are read as they are right now, including changes not yet stored.

From Lua (`LUA.CALL`, `LUA.ASYNC`, macros, streams), statements may also change data in the throwaway
in-memory database:

```lua
local rows = xl.sql("select name from [Orders] where qty > ?", {15})   -- table of rows, header first
local recs = xl.sql_records("select * from [Data!A1:B9] where k >= :min", {min = 10})  -- {{k=..., v=...}, ...}
xl.sql_exec("create temp table t(a); insert into t values (1), (2)")
xl.range("Out!A1"):set(rows)                                             -- write a result to the sheet (macro/async)
```

## The Lua ribbon tab

| Control | What it does |
|---|---|
| **Status button** | green: no scripts · yellow: trusted scripts running · red: scripts present but blocked or failed. Red: lists the scripts and offers to enable them. Yellow: offers to disable them again. |
| **Re-run async** | forgets the cached `LUA.ASYNC` results of the active workbook and runs those calls again |
| **Macros** | menu of the workbook's `xl.macro` functions; a **Stop running macros** item appears while any is running; progress and result go to the status bar, a failure opens a message box |
| **Sync folder** | links the workbook to a folder of scripts (next section) |
| **Files** | the Files panel |

### The Files panel

A panel docked on the right side of Excel's window (the workbook area gives up a strip) with two trees,
like a code explorer:

* **ThisWorkbook**: the files embedded in the open workbook. The panel follows the active workbook and edits
  it through `CustomXMLParts`; **save the workbook** to keep the changes (the footer says "Unsaved changes"
  until you do). Changes run at once: new functions and macros are available without reopening.
* **Machine**: the folder `%APPDATA%\xll-bridge\local` on this computer (shared Lua and native modules,
  machine databases; `require` and `xl.db.open` search it). It is real files on disk, needs no open
  workbook, and is created the first time you put something in it.

Seeing both side by side shows what travels with the workbook and what does not. **Drag and drop**
moves things between them and around:

| Drag | Within one tree | Between the trees |
|---|---|---|
| plain | **moves** | **copies** |
| with Ctrl | copies | |
| with Shift | | **moves** |

Drop on a folder (or on a file, meaning its folder, or on a root). Dragging files in from Explorer onto the
panel adds them to the selected folder (of either tree). A workbook that `require`s a machine module is no
longer self-contained: drag the module into ThisWorkbook to embed it.

Buttons (they act on the selected node; the two roots cannot be renamed or deleted):
**New folder** (name in place), **Add files…** / **Add folder…**, **Export…** (to anywhere on disk),
**Rename** (or F2), **Delete** (or the Delete key, with confirmation), **Copy over** (copies the selection
to the other tree's root), **Open in editor** (a machine file, in the program Windows associates with it,
Notepad if none). Embedded files are edited by [syncing a folder](#editing-scripts-in-your-own-editor) or by
copying them to Machine first.

Drag the panel's left edge to resize it. **Float** turns it into a tool window (**Dock** puts it back), ✕
closes it. There is deliberately **no script editor** in the panel.

If the workbook was green or yellow, files you add or change are trusted too (they are your own work).
A blocked workbook stays blocked until you enable it with the status button. Machine files are the user's
own and are not subject to that: a workbook's scripts are what is trusted, and they may `require` machine
modules once enabled.

## Editing scripts in your own editor

Keep your `.lua` files in a normal folder and edit them in VS Code (or anything else). **Lua ▸ Sync
folder** links the active, saved workbook to that folder:

* The folder is the source of truth: the workbook's embedded files become exactly the folder's files
  (dot-files such as `.git`, editor temp/backup files and `node_modules` are left out; text files get LF
  line endings). Edits made in the Files panel are overwritten by the next sync.
* Every change you save in the folder is copied into the workbook about two seconds later (once the
  folder has been quiet for a moment, so half-written files are never read); the new code starts running
  and the workbook recalculates. The status bar says what happened. The workbook is marked modified; save
  it to keep the files.
* Linking an **empty** folder to a workbook that already holds files offers to export them into the
  folder first, then follows it. A folder with no files never empties a workbook.
* Because you chose the folder, its files are trusted without asking again. The link is stored in
  `%APPDATA%\xll-bridge\sync.json` (created only when you link a folder), never in the workbook, so a
  workbook you receive cannot make your machine read a folder. Opening a linked workbook syncs it at once
  if the folder changed in the meantime.
* The synced files keep running even if something touches the workbook file on disk without changing its
  saved content (OneDrive, autosave); they are replaced only when the saved content really changes.
* Click the button again to sync immediately or to unlink (the files stay in the workbook).

## Trust model

* Scripts are read as XML only; **Lua is never created or compiled for an untrusted workbook.**
* Trust is stored in `%APPDATA%\xll-bridge\trust.json`, keyed by workbook path **and** a digest of all
  scripts. Editing any script outside the add-in, or moving the file, blocks the workbook again. A
  missing or corrupt trust file means nothing is trusted.
* A `LUA.*` call in a blocked workbook returns `#BLOCKED: …`. Enabling recalculates every open workbook so
  those cells fill in. Disabling (yellow button) deletes the workbook's trust entries, drops its Lua state
  and discards its cached background results and streams. (A macro already running is not interrupted by
  this; use **Stop running macros**.)
* `LUA.EVAL` in an unsaved workbook, or one with no scripts, runs in a **sandbox**: no
  `io`/`os`/`package`/`debug`/`load`/`dofile`, a 64 MB memory cap and an instruction limit. In a workbook
  whose scripts are blocked it stays blocked.
* Once enabled, scripts get the full Lua standard library (`io`, `os`, `package`, C modules). The trust
  prompt is the only barrier, which suits an internal tool where you write the scripts.

## Files on disk

Nothing is written at start-up; the `%APPDATA%\xll-bridge` folder appears only when one of these is
first needed:

* `trust.json`: trusted (workbook, script digest) pairs; written on the first Enable or panel edit
* `settings.json`: dock/float choice and panel width; written when you change them
* `sync.json`: workbook ↔ folder links; written when you link a folder
* `addin.log`: only when something went wrong (panic, crash, failed write), with the preceding events
* `local\`: the machine folder: shared Lua/DLL modules for `require` and machine databases; create it yourself (or let `xl.db.open(..., {create = "local"})` do)
* `%LOCALAPPDATA%\xll-bridge\cache\`: native modules extracted from a workbook, when it is enabled

The add-in also registers itself for the current user only (no admin rights): the COM classes for the
ribbon (`XllBridge.Ribbon`) and the stream server (`XllBridge.Rtd`) under `HKCU\Software\Classes`, and
the Excel add-in entry under `HKCU\Software\Microsoft\Office\Excel\Addins`. Removing the add-in from
the Add-ins dialog deletes those keys; `DllRegisterServer` / `DllUnregisterServer` are exported too.

## How it works

| Module | Role | Platform |
|---|---|---|
| `value` | Excel value model (`Value`, `Array`, `ErrorCode`) | any |
| `vfs` | the embedded file tree: folders, text/binary files, import/export, path safety | any |
| `package` | read/write that tree as a custom XML part (`urn:xll-bridge:files:v1`) in an `.xlsx`/`.xlsm` | any |
| `trust`, `runtime` | trust store and the per-workbook state machine (NoScripts / Blocked / Active / Failed) | any |
| `lua_host` | one Lua state per workbook, `Value` ⇄ Lua, the `xl` table | any |
| `xlapi` | cell access for Lua behind a trait; routes worker-thread requests to the UI thread | any |
| `jobs`, `streams`, `macros`, `procrun` | `LUA.ASYNC`, `LUA.STREAM`, macros, running programs | any |
| `sqldb`, `dbstore`, `dates` | SQLite over cells (`SQL.QUERY`, `xl.sql`), embedded and machine databases (`xl.db`), Excel date ⇄ ISO 8601 | any |
| `sync`, `panel`, `settings`, `ui`, `log` | folder sync, Files-panel commands, preferences, button states, diagnostics | any |
| `xll/` | XLOPER12 FFI (`MdCallBack12` is found at run time, no SDK), registration, exports | Windows x64 |
| `xll/ribbon/` | the COM add-in in the same DLL: ribbon, Files panel, cell access, RTD server, folder sync | Windows x64 |
| `bin/xllb-pack` | command-line tool that embeds `.lua` files into a workbook | any |

**Ribbon.** An XLL cannot add ribbon controls, so the same DLL is also an in-process COM add-in
(`IDTExtensibility2` + `IRibbonExtensibility`). `xlAutoOpen` registers it and connects it immediately, so
no Excel restart is needed. A 200 ms timer on Excel's UI thread follows workbook switches and drives
everything that has to happen there (async completion, cell requests from workers, macro results, RTD
updates, folder polling). If the COM side fails, the worksheet functions keep working.

**Finding the workbook of a formula.** A worksheet function cannot call back into Excel while it is
calculating, so the ribbon layer keeps a table of the open workbooks (name → real local path, from
`LocalFullName`) and the function looks its workbook up by name. `CELL("filename")` is only a fallback
for local files, because for synced files it returns a `https://` URL.

**The Files panel** is a child window of Excel's frame, not an Office task pane (that would need an ActiveX
host). It docks by shrinking the workbook area (`XLDESK`) and re-applying the strip every 50 ms; Excel's
window procedure is not subclassed. If Excel keeps undoing the resize the panel falls back to floating by
itself and notes why in `addin.log`.

**Streams** are served by an RTD server (`XllBridge.Rtd`) in the same DLL; Excel calls it by name through
`IDispatch`, which the server supports alongside the plain interface.

## Troubleshooting

* **`#NOPATH: …`**: the workbook could not be found; the text says what was seen, and `=LUA.STATUS()` prints
  the same. A newly opened workbook that contains embedded files is recalculated once when it appears.
* **`#BLOCKED: …`**: the scripts are not enabled yet; use the red status button.
* **A stream cell never shows a value**: `=LUA.STATUS()` reports how far the RTD server got (created /
  started / connected / disconnected, streams running now, and the last problem).
* **`#LUA: runtime error: Excel: …`** from a range call: the text after `Excel:` is Excel's own description.
* **`addin.log`** exists only after something went wrong; it records panics and access violations (with the
  faulting module and offset), preceded by the last events kept in memory. Attach it when reporting a crash.
* **"Excel is running into problems with the 'XllBridge.Ribbon' add-in"**: Excel disabled it. File ▸ Options
  ▸ Add-ins ▸ Manage: **Disabled Items** ▸ Go ▸ select it ▸ Enable, then restart Excel.
* **Testing new builds**: opening the `.xll` with File ▸ Open loads it for that session only. If Excel later
  complains about a missing add-in, remove the stale entries (File ▸ Options ▸ Add-ins, or delete
  `HKCU\Software\Microsoft\Office\Excel\Addins\XllBridge.Ribbon`).

## Limitations

* Functions are registered thread-unsafe, so recalculation is single-threaded.
* Scripts are read from the saved file, so re-save after embedding; a folder link runs the folder's files
  right away.
* The ribbon is untested against every Excel build; it uses 32×32 bitmap icons.
* Keystrokes typed into the docked panel (rename, new folder) depend on Excel passing them through; if that
  misbehaves on your Excel, use **Float**.
* x64 only.

## Future considerations

Recognised, but not on the roadmap:

* Streams that push a whole table (spilling over a range) instead of one value.
* `xl.http` and `xl.json` helpers, so background calls and streams can fetch and parse data without running
  an external program.

[mlua]: https://github.com/mlua-rs/mlua
