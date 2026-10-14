//! Excel XLL add-in exposing Lua for UDFs and data manipulation.
//!
//! The crate is split so everything except the Excel ABI can be built and tested anywhere:
//! `value`, `package` (script storage), `trust`, `lua_host` and `runtime` are portable;
//! `xll` (XLOPER12, registration, banner) is Windows-only.

pub mod lua_host;
pub mod package;
pub mod runtime;
pub mod trust;
pub mod ui;
pub mod value;

#[cfg(windows)]
mod xll;
