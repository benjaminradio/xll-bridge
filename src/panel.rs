//! Model for the "Lua files" panel: the commands the UI can issue and how the tree-only ones
//! are applied. Everything here is pure so it can be tested without Excel.

use std::path::PathBuf;

use crate::vfs::{Bundle, Kind, VfsError, clean_path};

// ------------------------------------------------------------------ commands from the UI

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Command {
    NewFolder {
        parent: String,
        name: String,
    },
    AddFiles {
        dest: String,
    },
    AddFolder {
        dest: String,
    },
    /// Files or folders dropped from Explorer.
    AddPaths {
        dest: String,
        paths: Vec<PathBuf>,
    },
    Export {
        path: String,
    },
    Delete {
        path: String,
    },
    Rename {
        path: String,
        name: String,
    },
    /// Copy (or move) a node onto a folder: between the workbook and the machine folder, or
    /// within one of them. Both are node ids (see [`MACHINE_ROOT`]).
    Transfer {
        from: String,
        to_dir: String,
        move_: bool,
    },
    /// Open a machine file in the user's editor.
    OpenInEditor {
        path: String,
    },
    Refresh,
    /// Switch between docked and floating.
    ToggleDock,
    /// Hide the panel.
    Close,
    /// Recreate the window in the currently configured mode (used by the docking fallback).
    Recreate,
}

// ------------------------------------------------------------------ node ids

/// The panel shows two trees: the workbook's embedded files ("ThisWorkbook") and the machine folder
/// ("Machine"). A node is identified by one string: a workbook node by its path in the file tree
/// (`""` is the ThisWorkbook root), a machine node by that path behind this marker (which is
/// never part of a valid file name), the machine root by the marker alone.
pub const MACHINE_ROOT: &str = "\u{1}";

pub fn is_machine(id: &str) -> bool {
    id.starts_with('\u{1}')
}

pub fn machine_id(rel: &str) -> String {
    format!("{MACHINE_ROOT}{rel}")
}

/// The path inside the machine folder (`""` for its root).
pub fn machine_rel(id: &str) -> &str {
    id.strip_prefix('\u{1}').unwrap_or(id)
}

pub fn is_root(id: &str) -> bool {
    id.is_empty() || id == MACHINE_ROOT
}

/// The node a node hangs from (`None` for the two roots).
pub fn parent_id(id: &str) -> Option<String> {
    if is_root(id) {
        return None;
    }
    let (prefix, rel) = if is_machine(id) {
        (MACHINE_ROOT, machine_rel(id))
    } else {
        ("", id)
    };
    Some(format!(
        "{prefix}{}",
        rel.rsplit_once('/').map_or("", |(p, _)| p)
    ))
}

/// The last path component.
pub fn node_name(id: &str) -> &str {
    let rel = if is_machine(id) { machine_rel(id) } else { id };
    rel.rsplit('/').next().unwrap_or(rel)
}

pub fn join(parent: &str, name: &str) -> String {
    if parent.is_empty() {
        name.to_owned()
    } else {
        format!("{parent}/{name}")
    }
}

fn single_name(name: &str) -> Result<&str, VfsError> {
    let n = name.trim();
    if n.contains('/') || n.contains('\\') {
        return Err(VfsError::InvalidPath(
            name.to_owned(),
            "a name cannot contain / or \\",
        ));
    }
    clean_path(n)?; // reuse the full validation (reserved names, characters, length)
    Ok(n)
}

/// Applies the commands that only touch the tree (no dialogs). Returns a status message.
pub fn apply_edit(bundle: &mut Bundle, cmd: &Command) -> Result<String, VfsError> {
    match cmd {
        Command::NewFolder { parent, name } => {
            let path = join(parent, single_name(name)?);
            if bundle.kind_of(&path).is_some() {
                return Err(VfsError::Exists(path));
            }
            bundle.mkdir(&path)?;
            Ok(format!("Created folder {path}"))
        }
        Command::Rename { path, name } => {
            let parent = path.rsplit_once('/').map_or("", |(p, _)| p);
            let target = join(parent, single_name(name)?);
            bundle.rename(path, &target)?;
            Ok(format!("Renamed to {target}"))
        }
        Command::Delete { path } => {
            let n = bundle.remove(path)?;
            Ok(format!(
                "Deleted {path} ({n} file{})",
                if n == 1 { "" } else { "s" }
            ))
        }
        _ => Ok(String::new()),
    }
}

/// How many files a delete will remove (for the confirmation prompt).
pub fn delete_summary(bundle: &Bundle, path: &str) -> String {
    match bundle.kind_of(path) {
        Some(Kind::Dir) => {
            let n = bundle
                .entries()
                .iter()
                .filter(|e| {
                    e.kind == Kind::File
                        && e.path.len() > path.len()
                        && e.path[..path.len()].eq_ignore_ascii_case(path)
                        && e.path.as_bytes()[path.len()] == b'/'
                })
                .count();
            format!(
                "the folder \"{path}\" and the {n} file{} inside it",
                if n == 1 { "" } else { "s" }
            )
        }
        _ => format!("\"{path}\""),
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn node_ids_tell_the_two_trees_apart() {
        use super::*;
        assert!(is_root("") && is_root(MACHINE_ROOT) && !is_root("a"));
        assert_eq!(parent_id("a/b.lua").as_deref(), Some("a"));
        assert_eq!(parent_id("a").as_deref(), Some(""));
        assert_eq!(parent_id(""), None);
        let m = machine_id("lib/x.lua");
        assert!(is_machine(&m) && !is_machine("lib/x.lua"));
        assert_eq!(machine_rel(&m), "lib/x.lua");
        assert_eq!(parent_id(&m), Some(machine_id("lib")));
        assert_eq!(
            parent_id(&machine_id("x.lua")).as_deref(),
            Some(MACHINE_ROOT)
        );
        assert_eq!(parent_id(MACHINE_ROOT), None);
        assert_eq!(node_name(&m), "x.lua");
        assert_eq!(node_name("a/b.lua"), "b.lua");
        // The marker can never be part of a real path.
        assert!(crate::vfs::clean_path(&m).is_err());
    }

    use super::*;
    use crate::package::testutil::bundle_of;

    #[test]
    fn edit_commands_validate_names_and_report() {
        let mut b = bundle_of(&[("lib/u.lua", "y")]);
        let r = apply_edit(
            &mut b,
            &Command::NewFolder {
                parent: "lib".into(),
                name: "sub".into(),
            },
        )
        .unwrap();
        assert!(r.contains("lib/sub"));
        assert!(
            apply_edit(
                &mut b,
                &Command::NewFolder {
                    parent: String::new(),
                    name: "a/b".into()
                }
            )
            .is_err()
        );
        assert!(
            apply_edit(
                &mut b,
                &Command::NewFolder {
                    parent: String::new(),
                    name: "lib".into()
                }
            )
            .is_err()
        );
        assert!(
            apply_edit(
                &mut b,
                &Command::NewFolder {
                    parent: String::new(),
                    name: "..".into()
                }
            )
            .is_err()
        );
        apply_edit(
            &mut b,
            &Command::Rename {
                path: "lib/u.lua".into(),
                name: "v.lua".into(),
            },
        )
        .unwrap();
        assert!(b.get("lib/v.lua").is_some());
        assert!(
            apply_edit(
                &mut b,
                &Command::Rename {
                    path: "lib/v.lua".into(),
                    name: "../x".into()
                }
            )
            .is_err()
        );
        assert!(delete_summary(&b, "lib").contains("1 file"));
        let r = apply_edit(&mut b, &Command::Delete { path: "lib".into() }).unwrap();
        assert!(r.contains("1 file") && b.entries().is_empty());
    }
}
