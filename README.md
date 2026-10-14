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
| `panel` | model of the Lua files task pane (HTML page, state, commands) | any |
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
  recalculates every open workbook so the `#BLOCKED` cells fill in.
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
```

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

### The Files panel

The **Files** button on the Lua tab opens a task pane on the right with the workbook's tree. You can
create folders, add files and folders (multi-select, recursive), rename, delete (with confirmation),
export a file or a folder to disk, and refresh. It edits the **open** workbook through Excel's
`CustomXMLParts`, so there is nothing to reload: **save the workbook to apply the changes** (the pane
says "Unsaved changes" until you do). A folder named `lib` you add is available to `require` after saving.

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

## Files

* `%APPDATA%\xll-bridge\trust.json` — trusted (workbook, script digest) pairs; created on first Enable
* `%APPDATA%\xll-bridge\lua\` — shared Lua/DLL modules for `require`; created when the add-in loads

## Known limitations / next steps

* Workbooks must be saved to a local path (OneDrive/SharePoint copies are resolved via `LocalFullName`
  where Excel provides it). Scripts are read from the file on disk, so re-save after embedding.
* Functions are registered thread-unsafe, so recalculation is single-threaded.
* The ribbon is untested against every Excel build; it targets Excel 2010+ (32x32 bitmap icons).
* The Files panel hosts the system web-browser control in an Office task pane; it needs `Shell.Explorer.2`
  (present on Windows 10/11). Drag-and-drop into the panel and an in-panel text editor are not done yet.
* Not yet implemented: reading arbitrary ranges/cells from Lua (`xl.range`), writing back, async/RTD
  for long-running external programs, a process-spawn helper with timeouts, a "revoke trust" action.
* x64 only.

[mlua]: https://github.com/mlua-rs/mlua
