//! The machine folder (`%APPDATA%\xll-bridge\local`) as seen by the Files panel: listing it and
//! changing it, with every path checked so nothing can escape the folder.
//!
//! Machine Lua modules, native modules and databases live here and are shared by every workbook.
//! The folder is created only when something is first put into it.

use std::path::{Path, PathBuf};

use crate::vfs::{self, Bundle, Entry, Kind};

pub type Result<T> = std::result::Result<T, String>;

/// Names that are never listed or copied (the same litter folder sync leaves out).
fn ignored(name: &str) -> bool {
    name.starts_with('.')
        || name.starts_with("~$")
        || name.ends_with('~')
        || name.ends_with(".tmp")
        || name.ends_with(".swp")
}

/// `rel` ("" for the root) as a path inside `root`, refusing anything that is not a plain
/// relative path or that passes through a symbolic link.
fn resolve(root: &Path, rel: &str) -> Result<PathBuf> {
    if rel.is_empty() {
        return Ok(root.to_path_buf());
    }
    let clean = vfs::clean_path(rel).map_err(|e| e.to_string())?;
    let mut path = root.to_path_buf();
    for part in clean.split('/') {
        path.push(part);
        if std::fs::symlink_metadata(&path).is_ok_and(|m| m.file_type().is_symlink()) {
            return Err(format!("'{rel}' goes through a symbolic link"));
        }
    }
    Ok(path)
}

/// Everything in the folder in tree order (folders before files, by name). A folder that does not
/// exist yet is simply empty.
pub fn list(root: &Path) -> Result<Vec<Entry>> {
    let mut out = Vec::new();
    walk(root, "", &mut out, 0)?;
    out.sort_by_key(|e| {
        let segs: Vec<&str> = e.path.split('/').collect();
        let last = segs.len() - 1;
        segs.iter()
            .enumerate()
            .map(|(i, s)| {
                (
                    u8::from(i == last && e.kind == Kind::File),
                    s.to_lowercase(),
                )
            })
            .collect::<Vec<_>>()
    });
    Ok(out)
}

fn walk(dir: &Path, rel: &str, out: &mut Vec<Entry>, depth: usize) -> Result<()> {
    if depth > 24 || out.len() > 5000 {
        return Ok(());
    }
    let Ok(rd) = std::fs::read_dir(dir) else {
        return Ok(()); // missing folder: empty
    };
    for e in rd.filter_map(|e| e.ok()) {
        let name = e.file_name();
        let Some(name) = name.to_str() else { continue };
        if ignored(name) {
            continue;
        }
        let Ok(meta) = std::fs::symlink_metadata(e.path()) else {
            continue;
        };
        if meta.file_type().is_symlink() {
            continue;
        }
        let path = if rel.is_empty() {
            name.to_owned()
        } else {
            format!("{rel}/{name}")
        };
        if meta.is_dir() {
            out.push(Entry {
                path: path.clone(),
                kind: Kind::Dir,
                size: 0,
            });
            walk(&e.path(), &path, out, depth + 1)?;
        } else {
            out.push(Entry {
                path,
                kind: Kind::File,
                size: meta.len() as usize,
            });
        }
    }
    Ok(())
}

pub fn new_folder(root: &Path, parent: &str, name: &str) -> Result<String> {
    let rel = crate::panel::join(parent, name.trim());
    let path = resolve(root, &rel)?;
    if path.exists() {
        return Err(format!("'{rel}' already exists"));
    }
    std::fs::create_dir_all(&path).map_err(|e| format!("{rel}: {e}"))?;
    Ok(format!("Created folder {rel} in the machine folder"))
}

pub fn rename(root: &Path, rel: &str, new_name: &str) -> Result<String> {
    if new_name.contains('/') || new_name.contains('\\') {
        return Err("a name cannot contain / or \\".into());
    }
    let from = resolve(root, rel)?;
    let parent = rel.rsplit_once('/').map_or("", |(p, _)| p);
    let target_rel = crate::panel::join(parent, new_name.trim());
    let to = resolve(root, &target_rel)?;
    if !from.exists() {
        return Err(format!("'{rel}' no longer exists"));
    }
    if to.exists() && !same_file_name_ignoring_case(&from, &to) {
        return Err(format!("'{target_rel}' already exists"));
    }
    std::fs::rename(&from, &to).map_err(|e| format!("{rel}: {e}"))?;
    Ok(format!("Renamed to {target_rel}"))
}

fn same_file_name_ignoring_case(a: &Path, b: &Path) -> bool {
    a.to_string_lossy().to_lowercase() == b.to_string_lossy().to_lowercase()
}

/// What a delete will remove, for the confirmation prompt.
pub fn delete_summary(root: &Path, rel: &str) -> String {
    let name = rel.rsplit('/').next().unwrap_or(rel);
    match resolve(root, rel) {
        Ok(p) if p.is_dir() => {
            let mut inside = Vec::new();
            let _ = walk(&p, "", &mut inside, 0);
            let files = inside.iter().filter(|e| e.kind == Kind::File).count();
            format!(
                "the folder '{name}' ({files} file{})",
                if files == 1 { "" } else { "s" }
            )
        }
        _ => format!("'{name}'"),
    }
}

pub fn delete(root: &Path, rel: &str) -> Result<String> {
    if rel.is_empty() {
        return Err("the machine folder itself cannot be deleted here".into());
    }
    let path = resolve(root, rel)?;
    let result = if path.is_dir() {
        std::fs::remove_dir_all(&path)
    } else {
        std::fs::remove_file(&path)
    };
    result.map_err(|e| format!("{rel}: {e}"))?;
    Ok(format!("Deleted {rel} from the machine folder"))
}

/// Copies files and folders from anywhere on disk into `dest` (a folder inside the machine folder).
pub fn add_paths(root: &Path, dest: &str, paths: &[PathBuf]) -> Result<String> {
    let dest_dir = resolve(root, dest)?;
    let (mut files, mut skipped) = (0usize, Vec::new());
    for p in paths {
        let Some(name) = p.file_name().and_then(|n| n.to_str()) else {
            skipped.push(format!("{}: no file name", p.display()));
            continue;
        };
        std::fs::create_dir_all(&dest_dir).map_err(|e| e.to_string())?;
        copy_rec(p, &dest_dir.join(name), &mut files, &mut skipped);
    }
    Ok(summary("Added", files, &skipped))
}

fn summary(verb: &str, files: usize, skipped: &[String]) -> String {
    let mut msg = format!("{verb} {files} file{}", if files == 1 { "" } else { "s" });
    if !skipped.is_empty() {
        msg.push_str(&format!(
            " \u{2014} skipped {}: {}",
            skipped.len(),
            skipped
                .iter()
                .take(2)
                .cloned()
                .collect::<Vec<_>>()
                .join("; ")
        ));
    }
    msg
}

fn copy_rec(from: &Path, to: &Path, files: &mut usize, skipped: &mut Vec<String>) {
    let Ok(meta) = std::fs::symlink_metadata(from) else {
        skipped.push(format!("{}: not readable", from.display()));
        return;
    };
    if meta.file_type().is_symlink() {
        skipped.push(format!("{}: symbolic link", from.display()));
    } else if meta.is_dir() {
        if let Err(e) = std::fs::create_dir_all(to) {
            skipped.push(format!("{}: {e}", to.display()));
            return;
        }
        let Ok(rd) = std::fs::read_dir(from) else {
            skipped.push(format!("{}: not readable", from.display()));
            return;
        };
        for e in rd.filter_map(|e| e.ok()) {
            if e.file_name().to_str().is_some_and(ignored) {
                continue;
            }
            copy_rec(&e.path(), &to.join(e.file_name()), files, skipped);
        }
    } else {
        match std::fs::copy(from, to) {
            Ok(_) => *files += 1,
            Err(e) => skipped.push(format!("{}: {e}", from.display())),
        }
    }
}

/// Copies a node of the workbook's file tree into the machine folder (`dest` is a folder in it).
pub fn copy_from_bundle(root: &Path, dest: &str, bundle: &Bundle, path: &str) -> Result<String> {
    let dest_dir = resolve(root, dest)?;
    std::fs::create_dir_all(&dest_dir).map_err(|e| e.to_string())?;
    let out = bundle
        .export_to(path, &dest_dir)
        .map_err(|e| e.to_string())?;
    Ok(format!("Copied to {}", out.display()))
}

/// Copies a machine node into the workbook's file tree (`dest` is a folder in it).
pub fn copy_to_bundle(root: &Path, rel: &str, bundle: &mut Bundle, dest: &str) -> Result<String> {
    let src = resolve(root, rel)?;
    if !src.exists() {
        return Err(format!("'{rel}' no longer exists"));
    }
    let r = bundle
        .import_path(&src, dest, true)
        .map_err(|e| e.to_string())?;
    Ok(summary("Copied", r.files, &r.skipped))
}

/// Copies or moves a node within the machine folder.
pub fn transfer_within(root: &Path, rel: &str, dest_dir: &str, move_: bool) -> Result<String> {
    let from = resolve(root, rel)?;
    let name = rel.rsplit('/').next().unwrap_or(rel);
    let target_rel = crate::panel::join(dest_dir, name);
    let to = resolve(root, &target_rel)?;
    if !from.exists() {
        return Err(format!("'{rel}' no longer exists"));
    }
    if to.starts_with(&from) {
        return Err("cannot put a folder into itself".into());
    }
    if from == to {
        return Ok(String::new());
    }
    if to.exists() {
        return Err(format!("'{target_rel}' already exists"));
    }
    if let Some(parent) = to.parent() {
        std::fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    if move_ {
        std::fs::rename(&from, &to).map_err(|e| format!("{rel}: {e}"))?;
        Ok(format!("Moved to {target_rel}"))
    } else {
        let (mut files, mut skipped) = (0, Vec::new());
        copy_rec(&from, &to, &mut files, &mut skipped);
        Ok(summary("Copied", files, &skipped))
    }
}

/// Copies a machine file to `dest` (a full file path chosen by the user).
pub fn export_file(root: &Path, rel: &str, dest: &Path) -> Result<String> {
    let src = resolve(root, rel)?;
    std::fs::copy(&src, dest).map_err(|e| format!("{rel}: {e}"))?;
    Ok(format!("Exported to {}", dest.display()))
}

/// Copies a machine folder into the existing folder `dest_parent`.
pub fn export_dir(root: &Path, rel: &str, dest_parent: &Path) -> Result<String> {
    let src = resolve(root, rel)?;
    let name = rel.rsplit('/').next().unwrap_or(rel);
    let (mut files, mut skipped) = (0, Vec::new());
    copy_rec(&src, &dest_parent.join(name), &mut files, &mut skipped);
    Ok(summary("Exported", files, &skipped))
}

/// Path of a machine file or folder (for opening it in an editor).
pub fn path_of(root: &Path, rel: &str) -> Result<PathBuf> {
    resolve(root, rel)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn names(root: &Path) -> Vec<String> {
        list(root).unwrap().into_iter().map(|e| e.path).collect()
    }

    #[test]
    fn listing_is_in_tree_order_and_skips_litter() {
        let d = tempfile::tempdir().unwrap();
        let r = d.path();
        for (p, c) in [
            ("b.lua", "1"),
            ("a/x.lua", "2"),
            ("a/.hidden", "3"),
            ("z~", "4"),
            ("a/sub/y.db", "5"),
        ] {
            let f = r.join(p);
            std::fs::create_dir_all(f.parent().unwrap()).unwrap();
            std::fs::write(f, c).unwrap();
        }
        assert_eq!(names(r), ["a", "a/sub", "a/sub/y.db", "a/x.lua", "b.lua"]);
        assert!(list(&r.join("missing")).unwrap().is_empty());
        assert!(
            !r.join("missing").exists(),
            "listing never creates the folder"
        );
    }

    #[test]
    fn folders_files_rename_and_delete() {
        let d = tempfile::tempdir().unwrap();
        let r = d.path().join("local");
        new_folder(&r, "", "lib").unwrap();
        assert!(new_folder(&r, "", "lib").is_err());
        let src = d.path().join("util.lua");
        std::fs::write(&src, "return {}").unwrap();
        assert!(
            add_paths(&r, "lib", &[src])
                .unwrap()
                .contains("Added 1 file")
        );
        assert_eq!(names(&r), ["lib", "lib/util.lua"]);
        rename(&r, "lib/util.lua", "tools.lua").unwrap();
        assert_eq!(names(&r), ["lib", "lib/tools.lua"]);
        assert!(rename(&r, "lib/tools.lua", "a/b").is_err());
        assert!(delete_summary(&r, "lib").contains("1 file"));
        delete(&r, "lib").unwrap();
        assert!(names(&r).is_empty());
        assert!(delete(&r, "").is_err());
    }

    #[test]
    fn paths_cannot_escape_the_folder() {
        let d = tempfile::tempdir().unwrap();
        let r = d.path().join("local");
        std::fs::create_dir_all(&r).unwrap();
        for bad in ["../x", "a/../../x", "C:/x", "a\\..\\b"] {
            assert!(resolve(&r, bad).is_err(), "{bad}");
        }
        // A leading slash is just ignored: the path still lands inside the folder.
        assert!(resolve(&r, "/etc/passwd").unwrap().starts_with(&r));
        assert!(new_folder(&r, "", "..").is_err());
        assert!(delete(&r, "../local").is_err());
    }

    #[cfg(unix)]
    #[test]
    fn symbolic_links_are_not_followed() {
        let d = tempfile::tempdir().unwrap();
        let (r, outside) = (d.path().join("local"), d.path().join("outside"));
        std::fs::create_dir_all(&r).unwrap();
        std::fs::create_dir_all(&outside).unwrap();
        std::fs::write(outside.join("secret.txt"), "x").unwrap();
        std::os::unix::fs::symlink(&outside, r.join("link")).unwrap();
        assert!(names(&r).is_empty());
        assert!(delete(&r, "link/secret.txt").is_err());
    }

    #[test]
    fn copying_between_the_workbook_and_the_machine_folder() {
        let d = tempfile::tempdir().unwrap();
        let r = d.path().join("local");
        let mut b = Bundle::new();
        b.put_file("lib/a.lua", b"a".to_vec(), false).unwrap();
        b.put_file("lib/data.db", b"db".to_vec(), false).unwrap();
        copy_from_bundle(&r, "", &b, "lib").unwrap();
        assert_eq!(names(&r), ["lib", "lib/a.lua", "lib/data.db"]);
        // And back, into another folder of the workbook.
        let mut b2 = Bundle::new();
        assert!(
            copy_to_bundle(&r, "lib", &mut b2, "from_machine")
                .unwrap()
                .contains("Copied 2")
        );
        assert_eq!(b2.get("from_machine/lib/a.lua").unwrap().data, b"a");
        // Within the machine folder: copy, move, and refusals.
        new_folder(&r, "", "archive").unwrap();
        transfer_within(&r, "lib/a.lua", "archive", false).unwrap();
        assert!(r.join("lib/a.lua").exists() && r.join("archive/a.lua").exists());
        transfer_within(&r, "lib/data.db", "archive", true).unwrap();
        assert!(!r.join("lib/data.db").exists() && r.join("archive/data.db").exists());
        assert!(transfer_within(&r, "lib", "lib", true).is_err());
        assert!(
            transfer_within(&r, "archive/a.lua", "lib", true).is_err(),
            "lib/a.lua exists"
        );
    }
}
