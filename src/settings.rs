//! The few user preferences, `%APPDATA%\xll-bridge\settings.json`. Read-only until the user
//! changes one; the folder is created only then.

use std::path::{Path, PathBuf};

use serde::{Deserialize, Serialize};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct Settings {
    /// Dock the Lua files panel into Excel's window (otherwise it floats).
    pub docked: bool,
    /// Docked panel width in pixels at 96 DPI.
    pub panel_width: i32,
}

impl Default for Settings {
    fn default() -> Self {
        Settings {
            docked: true,
            panel_width: 340,
        }
    }
}

pub fn path() -> Option<PathBuf> {
    crate::trust::TrustStore::default_path()
        .and_then(|p| p.parent().map(|d| d.join("settings.json")))
}

/// Missing or unreadable settings mean the defaults.
pub fn load_from(path: &Path) -> Settings {
    std::fs::read(path)
        .ok()
        .and_then(|b| serde_json::from_slice(&b).ok())
        .unwrap_or_default()
}

pub fn save_to(path: &Path, s: &Settings) -> std::io::Result<()> {
    if let Some(dir) = path.parent() {
        std::fs::create_dir_all(dir)?;
    }
    let json = serde_json::to_vec_pretty(s).map_err(std::io::Error::other)?;
    let tmp = path.with_extension("json.tmp");
    std::fs::write(&tmp, json)?;
    std::fs::rename(tmp, path)
}

pub fn load() -> Settings {
    path().map(|p| load_from(&p)).unwrap_or_default()
}

pub fn save(s: &Settings) {
    if let Some(p) = path() {
        let _ = save_to(&p, s);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_when_missing_or_corrupt_and_round_trip_when_saved() {
        let dir = tempfile::tempdir().unwrap();
        let file = dir.path().join("new").join("settings.json");
        assert_eq!(load_from(&file), Settings::default());
        assert!(
            !file.parent().unwrap().exists(),
            "reading never creates anything"
        );

        let s = Settings {
            docked: false,
            panel_width: 400,
        };
        save_to(&file, &s).unwrap();
        assert_eq!(load_from(&file), s);

        std::fs::write(&file, b"{ nope").unwrap();
        assert_eq!(load_from(&file), Settings::default());
        // Unknown/missing fields fall back per field.
        std::fs::write(&file, br#"{"docked": false}"#).unwrap();
        assert_eq!(
            load_from(&file),
            Settings {
                docked: false,
                panel_width: 340
            }
        );
    }
}
