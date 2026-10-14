//! The virtual file tree embedded in a workbook: folders, text files and binary files (DLLs, data).
//!
//! Paths are `/`-separated, relative, case-insensitive-unique (they must survive export to
//! Windows), and validated on every entry point so nothing like `..` can ever reach the disk.
//! Text files are stored with `\n` line endings.

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

use sha2::{Digest, Sha256};

pub const MAX_FILE_BYTES: usize = 32 * 1024 * 1024;
pub const MAX_TOTAL_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 16;
const MAX_SEGMENT: usize = 120;
const MAX_PATH: usize = 240;

#[derive(Debug, thiserror::Error)]
pub enum VfsError {
    #[error("invalid path '{0}': {1}")]
    InvalidPath(String, &'static str),
    #[error("'{0}' already exists")]
    Exists(String),
    #[error("'{0}' not found")]
    NotFound(String),
    #[error("{0}")]
    TooLarge(String),
    #[error("io: {0}")]
    Io(#[from] std::io::Error),
}

pub type Result<T> = std::result::Result<T, VfsError>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Kind {
    File,
    Dir,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct File {
    pub path: String,
    pub data: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Entry {
    pub path: String,
    pub kind: Kind,
    pub size: usize,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Bundle {
    files: Vec<File>,
    /// Explicitly created folders (folders that only exist because a file lives in them are implicit).
    dirs: Vec<String>,
}

#[derive(Debug, Default, PartialEq, Eq)]
pub struct ImportReport {
    pub files: usize,
    pub dirs: usize,
    pub replaced: usize,
    /// Human-readable reasons for anything left out.
    pub skipped: Vec<String>,
}

const RESERVED: [&str; 22] = [
    "con", "prn", "aux", "nul", "com1", "com2", "com3", "com4", "com5", "com6", "com7", "com8",
    "com9", "lpt1", "lpt2", "lpt3", "lpt4", "lpt5", "lpt6", "lpt7", "lpt8", "lpt9",
];

fn key(s: &str) -> String {
    s.to_lowercase()
}

/// Validates and canonicalises a path (`\` accepted as a separator; no leading/trailing slash).
pub fn clean_path(raw: &str) -> Result<String> {
    let bad = |why| VfsError::InvalidPath(raw.to_owned(), why);
    let s = raw.trim().replace('\\', "/");
    let s = s.trim_matches('/');
    if s.is_empty() {
        return Err(bad("empty"));
    }
    if s.len() > MAX_PATH {
        return Err(bad("too long"));
    }
    let segs: Vec<&str> = s.split('/').collect();
    if segs.len() > MAX_DEPTH {
        return Err(bad("too deep"));
    }
    for seg in &segs {
        if seg.is_empty() {
            return Err(bad("empty segment"));
        }
        if *seg == "." || *seg == ".." {
            return Err(bad("'.' and '..' are not allowed"));
        }
        if seg.len() > MAX_SEGMENT {
            return Err(bad("name too long"));
        }
        if seg
            .chars()
            .any(|c| c.is_control() || "<>:\"|?*".contains(c))
        {
            return Err(bad("contains a character Windows does not allow"));
        }
        if seg.ends_with('.') || seg.ends_with(' ') {
            return Err(bad("must not end with a dot or space"));
        }
        let stem = seg.split('.').next().unwrap_or("");
        if RESERVED.contains(&key(stem).as_str()) {
            return Err(bad("reserved device name"));
        }
    }
    Ok(segs.join("/"))
}

fn ancestors(path: &str) -> impl Iterator<Item = &str> {
    path.match_indices('/').map(move |(i, _)| &path[..i])
}

fn is_inside(path: &str, dir: &str) -> bool {
    path.len() > dir.len()
        && path.as_bytes()[dir.len()] == b'/'
        && key(&path[..dir.len()]) == key(dir)
}

fn same(a: &str, b: &str) -> bool {
    key(a) == key(b)
}

/// True if `data` can be stored as XML text without loss.
pub fn is_xml_text(data: &[u8]) -> bool {
    match std::str::from_utf8(data) {
        Ok(s) => s.chars().all(|c| {
            (c >= ' ' || c == '\t' || c == '\n')
                && c != '\u{feff}'
                && c != '\u{fffe}'
                && c != '\u{ffff}'
        }),
        Err(_) => false,
    }
}

fn normalize_text(data: Vec<u8>) -> Vec<u8> {
    if data.windows(2).any(|w| w == b"\r\n")
        && std::str::from_utf8(&data).is_ok()
        && !data.contains(&0)
    {
        let mut out = Vec::with_capacity(data.len());
        let mut i = 0;
        while i < data.len() {
            if data[i] == b'\r' && data.get(i + 1) == Some(&b'\n') {
                i += 1;
                continue;
            }
            out.push(data[i]);
            i += 1;
        }
        return out;
    }
    data
}

impl Bundle {
    pub fn new() -> Self {
        Bundle::default()
    }

    /// Keeps the internal order canonical so equal trees compare (and serialise) equal.
    fn canonicalise(&mut self) {
        self.files.sort_by(|a, b| a.path.cmp(&b.path));
        self.dirs.sort();
        self.dirs.dedup();
    }

    pub fn is_empty(&self) -> bool {
        self.files.is_empty() && self.dirs.is_empty()
    }

    pub fn file_count(&self) -> usize {
        self.files.len()
    }

    pub fn total_bytes(&self) -> usize {
        self.files.iter().map(|f| f.data.len()).sum()
    }

    pub fn files(&self) -> &[File] {
        &self.files
    }

    pub fn explicit_dirs(&self) -> &[String] {
        &self.dirs
    }

    pub fn kind_of(&self, path: &str) -> Option<Kind> {
        let p = clean_path(path).ok()?;
        if self.files.iter().any(|f| same(&f.path, &p)) {
            return Some(Kind::File);
        }
        let is_dir = self.dirs.iter().any(|d| same(d, &p))
            || self.files.iter().any(|f| is_inside(&f.path, &p))
            || self.dirs.iter().any(|d| is_inside(d, &p));
        is_dir.then_some(Kind::Dir)
    }

    pub fn get(&self, path: &str) -> Option<&File> {
        let p = clean_path(path).ok()?;
        self.files.iter().find(|f| same(&f.path, &p))
    }

    fn check_ancestors_free(&self, path: &str) -> Result<()> {
        for a in ancestors(path) {
            if self.files.iter().any(|f| same(&f.path, a)) {
                return Err(VfsError::InvalidPath(path.to_owned(), "a parent is a file"));
            }
        }
        Ok(())
    }

    /// Adds or replaces a file, creating folders implicitly.
    pub fn put_file(&mut self, path: &str, data: Vec<u8>, overwrite: bool) -> Result<bool> {
        let p = clean_path(path)?;
        let data = normalize_text(data);
        if data.len() > MAX_FILE_BYTES {
            return Err(VfsError::TooLarge(format!(
                "'{p}' exceeds the {} MB file limit",
                MAX_FILE_BYTES >> 20
            )));
        }
        self.check_ancestors_free(&p)?;
        if self.kind_of(&p) == Some(Kind::Dir) {
            return Err(VfsError::Exists(p));
        }
        let existing = self.files.iter().position(|f| same(&f.path, &p));
        let old = existing.map_or(0, |i| self.files[i].data.len());
        if self.total_bytes() - old + data.len() > MAX_TOTAL_BYTES {
            return Err(VfsError::TooLarge(format!(
                "workbook would exceed the {} MB limit for embedded files",
                MAX_TOTAL_BYTES >> 20
            )));
        }
        match existing {
            Some(_) if !overwrite => Err(VfsError::Exists(p)),
            Some(i) => {
                self.files[i] = File { path: p, data };
                self.canonicalise();
                Ok(true)
            }
            None => {
                self.files.push(File { path: p, data });
                self.canonicalise();
                Ok(false)
            }
        }
    }

    pub fn mkdir(&mut self, path: &str) -> Result<()> {
        let p = clean_path(path)?;
        self.check_ancestors_free(&p)?;
        match self.kind_of(&p) {
            Some(Kind::File) => Err(VfsError::Exists(p)),
            Some(Kind::Dir) => Ok(()),
            None => {
                self.dirs.push(p);
                self.canonicalise();
                Ok(())
            }
        }
    }

    /// Removes a file, or a folder and everything in it. Returns the number of files removed.
    pub fn remove(&mut self, path: &str) -> Result<usize> {
        let p = clean_path(path)?;
        if self.kind_of(&p).is_none() {
            return Err(VfsError::NotFound(p));
        }
        let before = self.files.len();
        self.files
            .retain(|f| !same(&f.path, &p) && !is_inside(&f.path, &p));
        self.dirs.retain(|d| !same(d, &p) && !is_inside(d, &p));
        Ok(before - self.files.len())
    }

    /// Renames or moves a file or folder. The target must not exist (a case-only change is fine).
    pub fn rename(&mut self, from: &str, to: &str) -> Result<()> {
        let (from, to) = (clean_path(from)?, clean_path(to)?);
        if self.kind_of(&from).is_none() {
            return Err(VfsError::NotFound(from));
        }
        if is_inside(&to, &from) {
            return Err(VfsError::InvalidPath(
                to,
                "cannot move a folder into itself",
            ));
        }
        if !same(&from, &to) && self.kind_of(&to).is_some() {
            return Err(VfsError::Exists(to));
        }
        self.check_ancestors_free(&to)?;
        let remap = |p: &str| -> Option<String> {
            if same(p, &from) {
                Some(to.clone())
            } else if is_inside(p, &from) {
                Some(format!("{to}{}", &p[from.len()..]))
            } else {
                None
            }
        };
        for f in &mut self.files {
            if let Some(n) = remap(&f.path) {
                f.path = n;
            }
        }
        for d in &mut self.dirs {
            if let Some(n) = remap(d) {
                *d = n;
            }
        }
        self.canonicalise();
        Ok(())
    }

    /// Every file and folder (including implicit folders), in tree order: a folder, then its
    /// contents, folders before files at each level, names case-insensitive.
    pub fn entries(&self) -> Vec<Entry> {
        let mut all: Vec<Entry> = Vec::new();
        let mut dirs: BTreeSet<String> = BTreeSet::new();
        for d in &self.dirs {
            dirs.insert(d.clone());
            dirs.extend(ancestors(d).map(str::to_owned));
        }
        for f in &self.files {
            dirs.extend(ancestors(&f.path).map(str::to_owned));
        }
        // Distinct by case-insensitive key.
        let mut seen = BTreeSet::new();
        for d in dirs {
            if seen.insert(key(&d)) {
                all.push(Entry {
                    path: d,
                    kind: Kind::Dir,
                    size: 0,
                });
            }
        }
        all.extend(self.files.iter().map(|f| Entry {
            path: f.path.clone(),
            kind: Kind::File,
            size: f.data.len(),
        }));
        let sort_key = |e: &Entry| -> Vec<(u8, String)> {
            let segs: Vec<&str> = e.path.split('/').collect();
            let last = segs.len() - 1;
            segs.iter()
                .enumerate()
                .map(|(i, s)| (u8::from(i == last && e.kind == Kind::File), key(s)))
                .collect()
        };
        all.sort_by_key(sort_key);
        all
    }

    /// Lua files in the root folder; these run (in name order) when the workbook is enabled.
    pub fn root_scripts(&self) -> Vec<&File> {
        let mut v: Vec<&File> = self
            .files
            .iter()
            .filter(|f| !f.path.contains('/') && key(&f.path).ends_with(".lua"))
            .collect();
        v.sort_by_key(|f| key(&f.path));
        v
    }

    /// Fingerprint of the whole tree; trust is pinned to this.
    pub fn digest(&self) -> String {
        let mut files: Vec<&File> = self.files.iter().collect();
        files.sort_by(|a, b| a.path.cmp(&b.path));
        let mut dirs: Vec<&String> = self.dirs.iter().collect();
        dirs.sort();
        let mut h = Sha256::new();
        for f in files {
            h.update(b"F");
            h.update((f.path.len() as u64).to_le_bytes());
            h.update(f.path.as_bytes());
            h.update((f.data.len() as u64).to_le_bytes());
            h.update(&f.data);
        }
        for d in dirs {
            h.update(b"D");
            h.update((d.len() as u64).to_le_bytes());
            h.update(d.as_bytes());
        }
        h.finalize().iter().map(|b| format!("{b:02x}")).collect()
    }

    // ------------------------------------------------------------ disk import / export

    /// Imports a file or folder from disk under `dest_dir` ("" for the root). Symlinks, unreadable
    /// and oversized items are skipped and reported rather than failing the whole import.
    pub fn import_path(
        &mut self,
        disk: &Path,
        dest_dir: &str,
        overwrite: bool,
    ) -> Result<ImportReport> {
        let mut report = ImportReport::default();
        let dest = if dest_dir.trim().is_empty() {
            String::new()
        } else {
            clean_path(dest_dir)?
        };
        let name = disk
            .file_name()
            .and_then(|n| n.to_str())
            .ok_or_else(|| VfsError::InvalidPath(disk.display().to_string(), "no file name"))?;
        let target = if dest.is_empty() {
            name.to_owned()
        } else {
            format!("{dest}/{name}")
        };
        self.import_rec(disk, &target, overwrite, &mut report)?;
        Ok(report)
    }

    fn import_rec(
        &mut self,
        disk: &Path,
        target: &str,
        overwrite: bool,
        rep: &mut ImportReport,
    ) -> Result<()> {
        let meta = match std::fs::symlink_metadata(disk) {
            Ok(m) => m,
            Err(e) => {
                rep.skipped.push(format!("{}: {e}", disk.display()));
                return Ok(());
            }
        };
        if meta.file_type().is_symlink() {
            rep.skipped
                .push(format!("{}: symbolic link", disk.display()));
        } else if meta.is_dir() {
            match clean_path(target).and_then(|t| self.mkdir(&t).map(|_| t)) {
                Ok(_) => rep.dirs += 1,
                Err(e) => {
                    rep.skipped.push(e.to_string());
                    return Ok(());
                }
            }
            let mut children: Vec<_> = std::fs::read_dir(disk)?.filter_map(|e| e.ok()).collect();
            children.sort_by_key(|e| e.file_name());
            for c in children {
                let n = c.file_name();
                let Some(n) = n.to_str() else {
                    rep.skipped
                        .push(format!("{}: name is not valid Unicode", c.path().display()));
                    continue;
                };
                self.import_rec(&c.path(), &format!("{target}/{n}"), overwrite, rep)?;
            }
        } else if meta.len() as usize > MAX_FILE_BYTES {
            rep.skipped.push(format!(
                "{}: larger than {} MB",
                disk.display(),
                MAX_FILE_BYTES >> 20
            ));
        } else {
            match std::fs::read(disk) {
                Err(e) => rep.skipped.push(format!("{}: {e}", disk.display())),
                Ok(data) => match self.put_file(target, data, overwrite) {
                    Ok(replaced) => {
                        rep.files += 1;
                        rep.replaced += usize::from(replaced);
                    }
                    Err(e) => rep.skipped.push(e.to_string()),
                },
            }
        }
        Ok(())
    }

    /// Writes `path` (file or folder) into the existing directory `dest_parent`, keeping its name.
    /// Existing files are overwritten. Returns the created path.
    pub fn export_to(&self, path: &str, dest_parent: &Path) -> Result<PathBuf> {
        let p = clean_path(path)?;
        let name = p.rsplit('/').next().unwrap_or(&p);
        let out = dest_parent.join(name);
        self.export_rec(&p, &out)?;
        Ok(out)
    }

    /// Writes a single file to an exact destination path.
    pub fn export_file_as(&self, path: &str, dest: &Path) -> Result<()> {
        let f = self
            .get(path)
            .ok_or_else(|| VfsError::NotFound(path.to_owned()))?;
        std::fs::write(dest, &f.data)?;
        Ok(())
    }

    fn export_rec(&self, path: &str, out: &Path) -> Result<()> {
        match self.kind_of(path) {
            None => Err(VfsError::NotFound(path.to_owned())),
            Some(Kind::File) => {
                std::fs::write(out, &self.get(path).expect("file exists").data)?;
                Ok(())
            }
            Some(Kind::Dir) => {
                std::fs::create_dir_all(out)?;
                let prefix_len = path.len() + 1;
                let mut direct: BTreeSet<(String, bool)> = BTreeSet::new();
                for e in self.entries() {
                    if is_inside(&e.path, path) {
                        let rest = &e.path[prefix_len..];
                        if !rest.contains('/') {
                            direct.insert((e.path.clone(), e.kind == Kind::Dir));
                        }
                    }
                }
                for (child, _) in direct {
                    let name = child.rsplit('/').next().unwrap_or(&child);
                    self.export_rec(&child, &out.join(name))?;
                }
                Ok(())
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn b() -> Bundle {
        let mut b = Bundle::new();
        b.put_file("main.lua", b"print(1)".to_vec(), false).unwrap();
        b.put_file("lib/util.lua", b"return {}".to_vec(), false)
            .unwrap();
        b.put_file("lib/deep/x.dat", vec![0, 1, 2, 255], false)
            .unwrap();
        b.mkdir("empty").unwrap();
        b
    }

    #[test]
    fn rejects_hostile_paths() {
        for bad in [
            "",
            "/",
            "..",
            "a/../b",
            "../x",
            "a//b",
            "C:/x",
            "a/b:c",
            "con",
            "a/NUL.txt",
            "a.",
            "x\u{1}y",
            "a/.",
            "a\\..\\b",
        ] {
            assert!(clean_path(bad).is_err(), "{bad:?} must be rejected");
        }
        assert_eq!(clean_path(" a\\b/c ").unwrap(), "a/b/c");
        assert_eq!(clean_path("/lib/").unwrap(), "lib");
        assert!(clean_path(&"a/".repeat(20)).is_err());
    }

    #[test]
    fn tree_order_is_folders_first_then_files() {
        let paths: Vec<String> = b().entries().into_iter().map(|e| e.path).collect();
        assert_eq!(
            paths,
            [
                "empty",
                "lib",
                "lib/deep",
                "lib/deep/x.dat",
                "lib/util.lua",
                "main.lua"
            ]
        );
    }

    #[test]
    fn case_insensitive_uniqueness_and_kinds() {
        let mut b = b();
        assert!(matches!(
            b.put_file("MAIN.LUA", vec![], false),
            Err(VfsError::Exists(_))
        ));
        assert!(b.put_file("MAIN.LUA", b"x".to_vec(), true).unwrap());
        assert_eq!(b.file_count(), 3);
        assert!(
            b.put_file("lib", vec![], true).is_err(),
            "cannot overwrite a folder with a file"
        );
        assert!(
            b.put_file("main.lua/child", vec![], true).is_err(),
            "cannot nest under a file"
        );
        assert!(b.mkdir("main.lua").is_err());
        assert_eq!(b.kind_of("LIB/Deep"), Some(Kind::Dir));
    }

    #[test]
    fn crlf_is_normalised_but_binary_is_not() {
        let mut b = Bundle::new();
        b.put_file("a.lua", b"x\r\ny\r\n".to_vec(), false).unwrap();
        b.put_file("b.bin", b"x\r\ny\0".to_vec(), false).unwrap();
        assert_eq!(b.get("a.lua").unwrap().data, b"x\ny\n");
        assert_eq!(b.get("b.bin").unwrap().data, b"x\r\ny\0");
    }

    #[test]
    fn remove_and_rename_move_subtrees() {
        let mut b = b();
        b.rename("lib", "src/lib").unwrap();
        assert!(b.get("src/lib/util.lua").is_some() && b.get("lib/util.lua").is_none());
        assert!(b.rename("src", "src/lib/inner").is_err(), "into itself");
        assert!(b.rename("main.lua", "empty").is_err(), "target exists");
        b.rename("main.lua", "MAIN.lua").unwrap(); // case-only
        assert_eq!(b.remove("src").unwrap(), 2);
        assert!(matches!(b.remove("src"), Err(VfsError::NotFound(_))));
        assert_eq!(b.entries().len(), 2); // empty/, MAIN.lua
    }

    #[test]
    fn root_scripts_are_root_lua_in_name_order() {
        let mut b = b();
        b.put_file("A.LUA", vec![], false).unwrap();
        b.put_file("notes.txt", vec![], false).unwrap();
        let names: Vec<&str> = b.root_scripts().iter().map(|f| f.path.as_str()).collect();
        assert_eq!(names, ["A.LUA", "main.lua"]);
    }

    #[test]
    fn digest_ignores_insertion_order_but_not_content() {
        let mut one = Bundle::new();
        one.put_file("a", b"1".to_vec(), false).unwrap();
        one.put_file("b", b"2".to_vec(), false).unwrap();
        let mut two = Bundle::new();
        two.put_file("b", b"2".to_vec(), false).unwrap();
        two.put_file("a", b"1".to_vec(), false).unwrap();
        assert_eq!(one.digest(), two.digest());
        two.put_file("a", b"9".to_vec(), true).unwrap();
        assert_ne!(one.digest(), two.digest());
        one.mkdir("d").unwrap();
        assert_ne!(one.digest(), two.digest());
    }

    #[test]
    fn size_limits_apply() {
        let mut b = Bundle::new();
        assert!(matches!(
            b.put_file("big", vec![1; MAX_FILE_BYTES + 1], false),
            Err(VfsError::TooLarge(_))
        ));
        for i in 0..2 {
            b.put_file(&format!("f{i}"), vec![1; MAX_FILE_BYTES], false)
                .unwrap();
        }
        assert!(matches!(
            b.put_file("f2", vec![1; MAX_FILE_BYTES], false),
            Err(VfsError::TooLarge(_))
        ));
    }

    #[test]
    fn import_and_export_round_trip_a_folder() {
        let tmp = tempfile::tempdir().unwrap();
        let src = tmp.path().join("proj");
        std::fs::create_dir_all(src.join("sub/inner")).unwrap();
        std::fs::write(src.join("a.lua"), "return 1\r\n").unwrap();
        std::fs::write(src.join("sub/b.bin"), [0u8, 159, 146, 150]).unwrap();

        let mut b = Bundle::new();
        let rep = b.import_path(&src, "", true).unwrap();
        assert_eq!((rep.files, rep.dirs), (2, 3));
        assert!(b.get("proj/a.lua").is_some());
        assert_eq!(b.kind_of("proj/sub/inner"), Some(Kind::Dir));

        let out = tmp.path().join("out");
        std::fs::create_dir(&out).unwrap();
        let written = b.export_to("proj", &out).unwrap();
        assert_eq!(std::fs::read(written.join("a.lua")).unwrap(), b"return 1\n");
        assert_eq!(
            std::fs::read(written.join("sub/b.bin")).unwrap(),
            [0u8, 159, 146, 150]
        );
        assert!(written.join("sub/inner").is_dir());

        // Re-importing the same folder replaces rather than duplicating.
        let rep = b.import_path(&src, "", true).unwrap();
        assert_eq!(rep.replaced, 2);
        assert_eq!(b.file_count(), 2);
    }

    #[test]
    fn import_into_a_destination_folder_and_single_file_export() {
        let tmp = tempfile::tempdir().unwrap();
        let f = tmp.path().join("m.lua");
        std::fs::write(&f, "x").unwrap();
        let mut b = Bundle::new();
        b.import_path(&f, "lib/sub", true).unwrap();
        assert!(b.get("lib/sub/m.lua").is_some());
        let dest = tmp.path().join("copy.lua");
        b.export_file_as("lib/sub/m.lua", &dest).unwrap();
        assert_eq!(std::fs::read(dest).unwrap(), b"x");
        assert!(b.import_path(&f, "../escape", true).is_err());
    }
}
