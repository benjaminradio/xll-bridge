//! Excel XLL add-in exposing Lua for UDFs and data manipulation.
//!
//! The crate is split so everything except the Excel ABI can be built and tested anywhere:
//! `value`, `package` (script storage), `trust`, `lua_host` and `runtime` are portable;
//! `xll` (XLOPER12, registration, banner) is Windows-only.

pub mod dock;
pub mod jobs;
pub mod log;
pub mod lua_host;
pub mod package;
pub mod panel;
pub mod procrun;
pub mod runtime;
pub mod settings;
pub mod trust;
pub mod ui;
pub mod value;
pub mod vfs;

#[cfg(windows)]
mod xll;
