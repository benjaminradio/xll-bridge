//! What the ribbon button shows for a workbook. Pure data so it can be tested anywhere.
//!
//! * green  — no Lua scripts in the workbook
//! * yellow — trusted scripts, Lua is running
//! * red    — scripts present but blocked (or failed to load); nothing has run

use crate::runtime::Status;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Light {
    Green,
    Yellow,
    Red,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct View {
    pub light: Light,
    pub label: String,
    pub screentip: String,
    pub supertip: String,
}

fn list(names: &[String]) -> String {
    names
        .iter()
        .map(|n| format!("\u{2022} {n}"))
        .collect::<Vec<_>>()
        .join("\n")
}

/// For an unsaved workbook, a non-local one (URL), or no workbook at all.
pub fn view_without_file() -> View {
    View {
        light: Light::Green,
        label: "Lua: none".into(),
        screentip: "No Lua scripts".into(),
        supertip: "This workbook is not saved to a local file, so it cannot contain Lua scripts. \
                   LUA.EVAL still works for simple expressions."
            .into(),
    }
}

pub fn view(status: &Status) -> View {
    match status {
        Status::NoScripts => View {
            light: Light::Green,
            label: "Lua: none".into(),
            screentip: "No Lua scripts".into(),
            supertip: "This workbook has no embedded Lua scripts.".into(),
        },
        Status::Active { names } => View {
            light: Light::Yellow,
            label: "Lua: running".into(),
            screentip: "Lua scripts are running".into(),
            supertip: format!("Trusted files loaded for this workbook:\n{}", list(names)),
        },
        Status::Blocked { names } => View {
            light: Light::Red,
            label: "Lua: blocked".into(),
            screentip: "Lua scripts are blocked".into(),
            supertip: format!(
                "This workbook contains embedded Lua files that have not been enabled. Nothing has been run.\n{}\n\nClick to review and enable.",
                list(names)
            ),
        },
        Status::Failed { message } => View {
            light: Light::Red,
            label: "Lua: error".into(),
            screentip: "Lua scripts failed to load".into(),
            supertip: message.clone(),
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn colours_follow_the_trust_state() {
        assert_eq!(view(&Status::NoScripts).light, Light::Green);
        assert_eq!(
            view(&Status::Active {
                names: vec!["a".into()]
            })
            .light,
            Light::Yellow
        );
        assert_eq!(
            view(&Status::Blocked {
                names: vec!["a".into()]
            })
            .light,
            Light::Red
        );
        assert_eq!(
            view(&Status::Failed {
                message: "x".into()
            })
            .light,
            Light::Red
        );
        assert_eq!(view_without_file().light, Light::Green);
    }

    #[test]
    fn blocked_tip_names_the_scripts() {
        let v = view(&Status::Blocked {
            names: vec!["main".into(), "util".into()],
        });
        assert!(v.supertip.contains("main") && v.supertip.contains("util"));
    }
}
