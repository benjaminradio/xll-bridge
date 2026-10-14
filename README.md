# xll-bridge

A native Excel add-in (XLL, Rust) that exposes Lua (5.4 via [mlua]) for UDFs and data manipulation.
Scripts live **inside the workbook** as a custom XML part and are **disabled until you enable them**,
like VBA macros.

> Status: scaffold. The portable core is unit-tested on Linux; the Windows/Excel layer
> cross-compiles and exports the expected symbols but **has not been run inside Excel yet**.

## Layout

| Module | Role | Platform |
|---|---|---|
| `value` | Excel value model (`Value`, `Array`, `ErrorCode`) | any |
| `package` | read/write the scripts custom XML part in an `.xlsx`/`.xlsm` | any |
| `trust` | trust store pinned to (workbook path, script SHA-256) | any |
| `lua_host` | Lua state per workbook, `Value` ⇄ Lua | any |
| `runtime` | per-workbook state machine: NoScripts / Blocked / Active / Failed | any |
| `xll/` | XLOPER12 FFI, registration, exports, security banner | Windows x64 |
| `bin/xllb-pack` | CLI that embeds `.lua` files into a workbook | any |

## Trust model

* Scripts are read as XML only; **Lua is never created or compiled for an untrusted workbook.**
* Trust is stored in `%APPDATA%\xll-bridge\trust.json`, keyed by workbook path **and** a digest of all
  scripts. Editing any script, or moving the file, brings the banner back. A missing/corrupt trust
  file means nothing is trusted.
* When a `LUA.*` call hits a blocked workbook it returns `#BLOCKED: …` and a yellow bar appears at
  the top of Excel with **Enable Lua** / ✕.
* Once enabled, scripts get the full Lua standard library (`io`, `os`, `package`, C modules).
  The trust prompt is the only barrier — appropriate for an internal tool where you write the scripts.

## Using it

```sh
cargo build --release --target x86_64-pc-windows-gnu      # or -msvc on Windows
cp target/x86_64-pc-windows-gnu/release/xll_bridge.dll xll_bridge.xll
cargo run --bin xllb-pack -- embed Book.xlsx Book-lua.xlsx main.lua
cargo run --bin xllb-pack -- list Book-lua.xlsx
```

Open the `.xll` in Excel (File ▸ Options ▸ Add-ins ▸ Go ▸ Browse), open the workbook, then:

```
=LUA.CALL("add", A1, B1)         calls global function add(a, b) from the embedded scripts
=LUA.EVAL("... * 2", A1)         evaluates an expression/chunk; args arrive as ...
=LUA.STATUS()                    none / disabled / enabled
```

Lua conventions: ranges are tables of rows (`{{1,2},{3,4}}`), a flat list returns a column, empty
cells are `nil` (`xl.EMPTY` inside arrays), errors are `xl.NA`, `xl.VALUE`, `xl.DIV0`, …
`require` searches the workbook's folder then `%APPDATA%\xll-bridge\lua` (`.lua` and `.dll`).

## Known limitations / next steps

* After clicking **Enable Lua**, press **Ctrl+Alt+F9**: UDF cells that returned `#BLOCKED` are not dirty,
  so Excel won't recompute them on its own.
* The bar is a window owned by Excel positioned under a fixed ribbon offset, not Office's real message
  bar (not available to XLLs). Layout may need tuning on your Excel build/DPI.
* Functions are registered thread-unsafe, so recalculation is single-threaded.
* Not yet implemented: reading arbitrary ranges/cells from Lua (`xl.range`), writing back, async/RTD
  for long-running external programs, a process-spawn helper with timeouts.
* x64 only.

[mlua]: https://github.com/mlua-rs/mlua

## CI and releases

* `.github/workflows/ci.yml` — on every push/PR: fmt, clippy (`-D warnings`) and tests on Linux; clippy,
  tests and a release build on Windows x64 (MSVC). The Windows job uploads an artifact
  **`xll-bridge-windows-x64`** containing `xll-bridge.xll`, `xllb-pack.exe`, a ready-to-open
  `demo.xlsx` with embedded Lua, `demo.lua`, and `SHA256SUMS.txt`.
* `.github/workflows/release.yml` — pushing a tag like `v0.1.0` runs the same build and publishes a
  GitHub release with a zip of that artifact.

To test: download the artifact, open Excel, load `xll-bridge.xll` (File ▸ Options ▸ Add-ins ▸ Go ▸
Browse; you may need to right-click the file ▸ Properties ▸ Unblock), then open `demo.xlsx`.
