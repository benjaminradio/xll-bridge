# xll-bridge

A native Excel add-in (XLL, Rust) that exposes Lua (5.4 via [mlua]) for UDFs and data manipulation.
Files (Lua scripts, modules, data, even DLLs) live **inside the workbook** as a custom XML part and are **disabled until you enable them**,
like VBA macros. A ribbon button shows the state at a glance.

> Status: scaffold. The portable core is unit-tested on Linux; the Windows/Excel layer
> cross-compiles and exports the expected symbols but **has not been run inside Excel yet**.

## Layout

| Module | Role | Platform |
|---|---|---|
| `value` | Excel value model (`Value`, `Array`, `ErrorCode`) | any |
| `vfs` | the embedded file tree: folders, text/binary files, import/export, path safety | any |
| `package` | read/write that tree as a custom XML part in an `.xlsx`/`.xlsm` | any |
| `panel` | commands of the Lua files panel and how tree edits are applied | any |
| `log` | tiny diagnostic log (`%APPDATA%\xll-bridge\addin.log`) | any |
| `trust` | trust store pinned to (workbook path, script SHA-256) | any |
| `lua_host` | Lua state per workbook, `Value` ⇄ Lua | any |
| `runtime` | per-workbook state machine: NoScripts / Blocked / Active / Failed | any |
| `ui` | what the ribbon button shows for each state (colour, label, tooltips) | any |
| `xll/` | XLOPER12 FFI, registration, exports | Windows x64 |
| `xll/ribbon/` | COM add-in (same DLL): Lua ribbon tab, status button, Lua files task pane | Windows x64 |
| `bin/xllb-pack` | CLI that embeds `.lua` files into a workbook | any |

## Trust model

* Scripts are read as XML only; **Lua is never created or compiled for an untrusted workbook.**
* Trust is stored in `%APPDATA%\xll-bridge\trust.json`, keyed by workbook path **and** a digest of all
  scripts. Editing any script, or moving the file, brings the banner back. A missing/corrupt trust
  file means nothing is trusted.
* When a `LUA.*` call hits a blocked workbook it returns `#BLOCKED: …`. The **Lua** ribbon tab has one
  button that reflects the active workbook:

  | colour | meaning |
  |---|---|
  | green  | no Lua scripts in the workbook |
  | yellow | trusted scripts are loaded and running |
  | red    | scripts present but **blocked** (or failed to load) — nothing has run |

  Clicking the red button lists the scripts and asks for confirmation; **Yes** trusts them and
  recalculates every open workbook so the `#BLOCKED` cells fill in. Clicking the yellow button
  offers to **disable** the scripts again: the workbook's trust entries are deleted, its Lua state
  is dropped and its cached background results are discarded. (A macro already running is not
  interrupted; it stops at its 5 minute limit.)
* `LUA.EVAL` in an unsaved workbook, or one with no scripts, runs in a **sandbox**: no `io`/`os`/`package`/
  `debug`/`load`/`dofile`, 64 MB memory cap and an instruction limit. In a workbook whose scripts are
  blocked it stays blocked.
* Once enabled, scripts get the full Lua standard library (`io`, `os`, `package`, C modules).
  The trust prompt is the only barrier — appropriate for an internal tool where you write the scripts.

## Using it

```sh
cargo build --release --target x86_64-pc-windows-gnu      # or -msvc on Windows
cp target/x86_64-pc-windows-gnu/release/xll_bridge.dll xll_bridge.xll
cargo run --bin xllb-pack -- embed Book.xlsx Book-lua.xlsx main.lua lib/   # files and folders
cargo run --bin xllb-pack -- list Book-lua.xlsx
cargo run --bin xllb-pack -- extract Book-lua.xlsx ./out
```

Open the `.xll` in Excel (File ▸ Options ▸ Add-ins ▸ Go ▸ Browse), open the workbook, then:

```
=LUA.CALL("add", A1, B1)         calls global function add(a, b) from the embedded scripts
=LUA.EVAL("... * 2", A1)         evaluates an expression/chunk; args arrive as ...
=LUA.STATUS()                    none / disabled / enabled
=LUA.ASYNC("slow", A1)           like LUA.CALL, but in the background (see below)
```

### Slow calls: `LUA.ASYNC`

`LUA.CALL` runs on Excel's UI thread, so a slow function (a web request, an external program) freezes
Excel while it runs. `=LUA.ASYNC("name", args…)` runs it on a background worker instead: the cell shows
`#GETTING_DATA` (what Excel's own async functions show) and fills in when the result arrives, with only
the waiting cells recalculated. Excel stays usable throughout.

* Each worker has its **own Lua state**, built from the workbook's enabled files: globals set by one call
  are not visible to `LUA.CALL` or to other workers. Pass what a function needs as arguments.
* Results are **cached** per function, arguments and script version, so recalculating does not re-run a
  slow call. They are re-run when the arguments or scripts change, or with **Re-run async** on the Lua tab.
* Every call has a 120 s limit and is cancelled when re-run or when the workbook closes (a Lua loop is
  interrupted; an external program started with `xl.run` is killed at its own timeout).
* In manual calculation mode finished cells are marked dirty and appear on the next F9.
* Blocked/untrusted workbooks behave as for `LUA.CALL`: nothing runs.

### Reading and writing cells from Lua

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

Addresses may be `Sheet!A1`, `'My Sheet'!A1`, `[Other.xlsx]Sheet!A1` or a defined name. Empty cells
read as `xl.EMPTY` inside tables; errors read as `xl.NA`, `xl.DIV0`, ... Multi-area ranges are
not supported and a single operation is limited to 5,000,000 cells.

* **`LUA.CALL` can read but not write** (Excel does not let a worksheet function change cells).
  Reads inside a function are not dependency-tracked: pass the range as an argument if the result
  must recalculate when it changes.
* **`LUA.ASYNC` can read and write.** The Lua code runs on a worker thread; each range operation is
  handed to Excel's UI thread and the worker waits for the answer, so batch your reads and writes
  (one `values()`/`set()` of a block, not one call per cell). If Excel is busy (editing a cell,
  a dialog) the operation is retried for a while, then fails with an error.
* There is deliberately no script editor in the panel: edit your `.lua` files in your own editor
  and use *Import* (or drag them onto the panel).

### Live data: `LUA.STREAM`

```lua
function clock()                 -- =LUA.STREAM("clock")
  while true do
    xl.emit(os.date("%H:%M:%S")) -- push a value; the cell updates
    xl.sleep(1)                  -- wait without blocking Excel (ends promptly when cancelled)
  end
end
```

A stream function runs on its own thread for as long as some cell shows it, and calls
`xl.emit(value)` to update the cells (single values only; a value *returned* at the end is pushed
too). Cells with the same function and arguments share one stream. It is implemented with Excel's
RTD mechanism (the DLL is also registered as the COM server `XllBridge.Rtd`, under HKCU like the
ribbon add-in). Only the latest value counts, so a fast stream cannot build a backlog; Excel is
told about new values by a timer, but Excel itself batches real-time updates to its *throttle interval* (2 seconds by default, so a stream that emits every second shows a new value every 2 seconds; lower it with `Application.RTD.ThrottleInterval = 500` in the VBA immediate window, a per-user Excel setting). The stream stops when no
cell uses it, when the workbook's scripts are disabled, or when Excel exits. An error ends it and
shows `#LUA: ...`. If a stream cell never shows a value, `=LUA.STATUS()` reports how far the RTD server got (created / started / connected / disconnected, streams running now, and the last problem). Streams can read and write cells with `xl.range` like `LUA.ASYNC` calls.

### Macros

`xl.macro("Name", function() ... end)` registers a function that appears in the ribbon's
**Lua > Macros** menu (rebuilt each time the menu opens). Clicking it runs the function in the
background, in a fresh Lua state built from the workbook's trusted files, so Excel stays
responsive. Unlike `LUA.CALL`, a macro may change cells. The status bar shows progress and the
result; a failure opens a message box. A macro is stopped after 5 minutes and the same macro
cannot run twice at once in a workbook. While any macro is running the menu has a **Stop running macros** item; closing Excel stops macros, streams and background calls, and a program started with `xl.run` is killed with them.

### Running programs from Lua

```lua
local r = xl.run{ "git", "status", "--short", cwd = "C:\\work\\repo", timeout = 10 }
-- r = { code, stdout, stderr, ok, timed_out, truncated } ; or nil, "message" if it could not start
local s = xl.shell("dir /b | find \".lua\"", { timeout = 5 })   -- through cmd /C (sh -c elsewhere)
```

Output is captured (32 MB cap per stream), `stdin = "text"` is supported, and `timeout` (seconds, default 60)
really kills the process (and its children on Windows). No console window appears. These exist only in
enabled workbooks; the sandbox used by `LUA.EVAL` in other workbooks has no way to start programs.

### The embedded file tree

* `.lua` files in the **root** run (in name order) when the workbook is enabled and define the globals
  `LUA.CALL` can call.
* Files in **folders** are modules: `require("lib.util")` loads `lib/util.lua` (or `lib/util/init.lua`)
  from the workbook.
* `xl.read_file("data/x.csv")` returns any embedded file's bytes as a Lua string; `xl.files()` lists them.
* `.dll`/`.so` files are extracted (only after the workbook is enabled) to
  `%LOCALAPPDATA%\xll-bridge\cache\<digest>\` and added to `package.cpath`, so `require("native.thing")`
  finds `native/thing.dll`.
* Paths are validated everywhere (no `..`, drive letters, reserved Windows names); text files are stored
  with `\n` line endings; limits are 32 MB per file and 64 MB per workbook.

### Editing scripts in your own editor: Sync folder

Keep your `.lua` files in a normal folder and edit them in VS Code (or anything else). **Lua >
Sync folder** links the active, saved workbook to that folder:

* The folder is the source of truth: the workbook's embedded files become exactly the folder's files
  (dot-files such as `.git`, editor temp/backup files and `node_modules` are left out, text files get
  LF line endings). Edits made in the Files panel are overwritten by the next sync.
* Every change you save in the folder is copied into the workbook about two seconds later (once the
  folder has been quiet for a moment, so half-written files are never read), the new code starts
  running and the workbook recalculates. The status bar says what happened. The workbook is marked
  modified; save it to keep the files.
* Because you chose the folder, its files are trusted without asking again. The link lives in
  `%APPDATA%\xll-bridge\sync.json` (created only when you link a folder), never in the workbook, so
  a workbook you receive cannot make your machine read a folder. Opening a linked workbook syncs it
  at once if the folder changed in the meantime.
* Click the button again to sync immediately or to unlink (the files stay in the workbook).
* A folder with no files never empties a workbook.

### The Files panel

The **Files** button on the Lua tab opens a panel docked on the right side of Excel, inside its
window: the workbook area gives up a strip and the panel sits in it, showing the workbook's tree.
Drag its left edge to resize it (the width is remembered). **Float** in the panel's header turns it into
a free-standing tool window instead (and **Dock** puts it back); ✕ closes it. If Excel will not let the
workbook area be resized (it undoes the change every tick) the panel falls back to floating by itself
and notes why in `addin.log`.

Docking works by watching Excel's workbook area (`XLDESK`, a child of the frame) every 50 ms and
re-applying the strip whenever Excel re-lays it out — Excel's window procedure is not subclassed. With
several Excel windows (one per workbook) it follows the window you are working in.

In the panel you can:

* **New folder** (type the name in place), **Add files…** / **Add folder…** (multi-select, folders are
  imported recursively), or **drag files and folders in from Explorer** onto the panel;
* **Rename** (button or F2, edited in place), **Delete** (button or Delete key, with confirmation);
* **Export…** a file (Save As) or a folder (pick a destination).

Additions go into the selected folder (or the selected file's folder). The panel follows the active
workbook and edits the **open** workbook through Excel's `CustomXMLParts`, so nothing needs reloading:
**save the workbook to apply the changes** (the footer says "Unsaved changes" until you do). A folder
named `lib` you add is available to `require` after saving.

Trust and edits: if the workbook was green or yellow, files you add or change are trusted too (they are
your own work), so they stay enabled after saving. A red/blocked workbook stays blocked — review the
files in the panel, then use the ribbon button. Files authored in an unsaved workbook are trusted once
saved.

Lua conventions: ranges are tables of rows (`{{1,2},{3,4}}`), a flat list returns a column, empty
cells are `nil` (`xl.EMPTY` inside arrays), errors are `xl.NA`, `xl.VALUE`, `xl.DIV0`, …
`require` searches the workbook's folder then `%APPDATA%\xll-bridge\lua` (`.lua` and `.dll`).

## Ribbon button: how it works

An XLL can't add ribbon controls, so the same DLL is also an in-process COM add-in
(`IDTExtensibility2` + `IRibbonExtensibility`). On load, `xlAutoOpen` registers it under
`HKCU\Software\Classes` and `HKCU\Software\Microsoft\Office\Excel\Addins\XllBridge.Ribbon` (no admin
rights) and then connects it immediately, so no Excel restart is needed. Removing the add-in from the
Add-ins dialog (`xlAutoRemove`) deletes those keys; `regsvr32`-style `DllRegisterServer` /
`DllUnregisterServer` are exported too. A 500 ms timer on Excel's UI thread notices workbook switches and
refreshes the button. If the COM side fails for any reason the worksheet functions keep working.

## Troubleshooting

* **How a formula finds its workbook:** a worksheet function cannot call back into Excel while it is
  calculating, so the ribbon layer keeps a table of the open workbooks (name -> real local path, via
  Excel's `LocalFullName`, which resolves OneDrive/SharePoint-synced files) and the function looks its
  workbook up by name. `CELL("filename")` is only a fallback for local files, because for synced files it
  returns a `https://` URL. If a workbook cannot be found you get `#NOPATH: …` with what was seen, and
  `=LUA.STATUS()` prints the same detail. A workbook that is newly opened and contains embedded files is
  recalculated once when it appears, so cells that ran too early fill in by themselves.

* `%APPDATA%\xll-bridge\addin.log` exists only after something went wrong. It records panics and access
  violations (with the faulting module and offset), preceded by the last events (start-up, button clicks)
  kept in memory. Attach it when reporting a crash.
* If Excel shows "Excel is running into problems with the 'XllBridge.Ribbon' add-in" it disables the add-in:
  File ▸ Options ▸ Add-ins ▸ Manage: **Disabled Items** ▸ Go ▸ select it ▸ Enable, then restart Excel.

## Files on disk (created only when needed)

Nothing is written at start-up; the `%APPDATA%\xll-bridge` folder appears only when one of these is first saved:

* `trust.json` — trusted (workbook, script digest) pairs; written on the first Enable or panel edit
* `settings.json` — dock/float choice and panel width; written when you change them
* `addin.log` — only when something went wrong (panic, crash, failed write), with the preceding events
* `lua\` — shared Lua/DLL modules for `require`; create it yourself if you want one
* `%LOCALAPPDATA%\xll-bridge\cache\` — native modules extracted from a workbook, when it is enabled

## Known limitations

* Workbooks must be saved to a local path (OneDrive/SharePoint copies are resolved via `LocalFullName`
  where Excel provides it). Scripts are read from the file on disk, so re-save after embedding; a
  folder link (see *Sync folder*) runs the folder's files right away.
* Functions are registered thread-unsafe, so recalculation is single-threaded.
* The ribbon is untested against every Excel build; it targets Excel 2010+ (32x32 bitmap icons).
* The docked panel is a child window of Excel's frame, not an Office task pane (that would need an
  ActiveX host). Keystrokes typed into it (rename, new folder) depend on Excel passing them through; if
  that misbehaves on your Excel, use **Float**.
* There is deliberately no script editor in the panel: edit the files in your own editor (use
  *Sync folder*).
* x64 only.

## Future considerations (recognised, not on the roadmap)

* Streams that push a whole table (spilling over a range) instead of one value.
* `xl.http` and `xl.json` helpers, so background calls and streams can fetch and parse data without
  running an external program.

[mlua]: https://github.com/mlua-rs/mlua
