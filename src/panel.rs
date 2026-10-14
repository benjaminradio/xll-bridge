//! Model for the "Lua files" task pane: the HTML page, the state it is sent, and the commands it
//! sends back. Everything here is pure so it can be tested without Excel.
//!
//! The pane is the built-in browser control hosted in an Office task pane. The host talks to it
//! through two functions of the page: it calls `setState({...})` to render, and polls `pull()`
//! for queued user commands (`name|arg|arg`, every part percent-encoded).

use crate::ui::{Light, View};
use crate::vfs::{Bundle, Kind, VfsError, clean_path};

pub const HTML: &str = include_str!("panel.html");

// ------------------------------------------------------------------ commands from the page

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Command {
    NewFolder { parent: String, name: String },
    AddFiles { dest: String },
    AddFolder { dest: String },
    Export { path: String },
    Delete { path: String },
    Rename { path: String, name: String },
    Refresh,
}

pub fn percent_decode(s: &str) -> String {
    let b = s.as_bytes();
    let mut out = Vec::with_capacity(b.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'%'
            && i + 2 < b.len()
            && let (Some(h), Some(l)) = (
                (b[i + 1] as char).to_digit(16),
                (b[i + 2] as char).to_digit(16),
            )
        {
            out.push((h * 16 + l) as u8);
            i += 3;
            continue;
        }
        out.push(b[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

pub fn parse_command(line: &str) -> Option<Command> {
    let parts: Vec<String> = line.trim().split('|').map(percent_decode).collect();
    let arg = |i: usize| parts.get(i).cloned().unwrap_or_default();
    Some(match parts.first()?.as_str() {
        "newfolder" => Command::NewFolder {
            parent: arg(1),
            name: arg(2),
        },
        "addfiles" => Command::AddFiles { dest: arg(1) },
        "addfolder" => Command::AddFolder { dest: arg(1) },
        "export" => Command::Export { path: arg(1) },
        "delete" => Command::Delete { path: arg(1) },
        "rename" => Command::Rename {
            path: arg(1),
            name: arg(2),
        },
        "refresh" => Command::Refresh,
        _ => return None,
    })
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

// ------------------------------------------------------------------ state sent to the page

/// A JavaScript string literal that is pure ASCII (so it survives any code page) and cannot break
/// out of a `<script>` or a string: control characters, quotes, `<`, `>` and `&` are escaped.
pub fn js_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (' '..='~').contains(&c) && !matches!(c, '<' | '>' | '&' | '\'') => out.push(c),
            c => {
                let mut buf = [0u16; 2];
                for unit in c.encode_utf16(&mut buf) {
                    out.push_str(&format!("\\u{unit:04x}"));
                }
            }
        }
    }
    out.push('"');
    out
}

pub struct PanelState<'a> {
    pub workbook: Option<&'a str>,
    /// False when the workbook can't be edited (no workbook, read-only/protected view, ...).
    pub ok: bool,
    pub hint: &'a str,
    pub view: &'a View,
    /// The live files differ from what is saved on disk.
    pub dirty: bool,
    /// Switching workbooks: tell the page to reset its selection and expansion.
    pub fresh: bool,
    /// (is_error, text)
    pub message: Option<(bool, String)>,
    pub bundle: Option<&'a Bundle>,
}

fn light_name(l: Light) -> &'static str {
    match l {
        Light::Green => "green",
        Light::Yellow => "yellow",
        Light::Red => "red",
    }
}

/// `setState({...})` call text for `execScript`.
pub fn state_js(s: &PanelState<'_>) -> String {
    let mut o = String::from("setState({");
    o.push_str(&format!(
        "wb:{},",
        s.workbook.map_or("null".to_owned(), js_string)
    ));
    o.push_str(&format!("ok:{},", s.ok));
    o.push_str(&format!("hint:{},", js_string(s.hint)));
    o.push_str(&format!("light:{},", js_string(light_name(s.view.light))));
    o.push_str(&format!("status:{},", js_string(&s.view.screentip)));
    o.push_str(&format!("dirty:{},fresh:{},", s.dirty, s.fresh));
    match &s.message {
        Some((err, text)) => o.push_str(&format!("msg:{},err:{},", js_string(text), err)),
        None => o.push_str("msg:\"\",err:false,"),
    }
    let entries = s.bundle.map(|b| b.entries()).unwrap_or_default();
    let n = entries.iter().filter(|e| e.kind == Kind::File).count();
    o.push_str(&format!("n:{n},entries:["));
    let roots: Vec<String> = s
        .bundle
        .map(|b| b.root_scripts().iter().map(|f| f.path.clone()).collect())
        .unwrap_or_default();
    for (i, e) in entries.iter().enumerate() {
        if i > 0 {
            o.push(',');
        }
        let kind = if e.kind == Kind::Dir { "d" } else { "f" };
        let run = u8::from(roots.contains(&e.path));
        o.push_str(&format!(
            "{{p:{},k:\"{kind}\",s:{},r:{run}}}",
            js_string(&e.path),
            e.size
        ));
    }
    o.push_str("]});");
    o
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::testutil::bundle_of;
    use crate::runtime::Status;

    #[test]
    fn page_exposes_the_functions_the_host_calls() {
        for needle in [
            "function pull()",
            "function setState(",
            "send(\"addfiles\"",
            "send(\"newfolder\"",
            "send(\"rename\"",
            "send(\"delete\"",
            "send(\"export\"",
            "send(\"refresh\"",
        ] {
            assert!(HTML.contains(needle), "panel.html is missing {needle}");
        }
        // Old embedded-browser document modes: no arrow functions, let/const, template strings.
        assert!(
            !HTML.contains("=>")
                && !HTML.contains("`")
                && !HTML.contains("const ")
                && !HTML.contains("let ")
        );
    }

    #[test]
    fn every_command_the_page_can_send_parses() {
        let cases = [
            (
                "newfolder|lib|new%20dir",
                Command::NewFolder {
                    parent: "lib".into(),
                    name: "new dir".into(),
                },
            ),
            (
                "addfiles|",
                Command::AddFiles {
                    dest: String::new(),
                },
            ),
            ("addfolder|a%2Fb", Command::AddFolder { dest: "a/b".into() }),
            (
                "export|x%2Fy.lua",
                Command::Export {
                    path: "x/y.lua".into(),
                },
            ),
            ("delete|x", Command::Delete { path: "x".into() }),
            (
                "rename|a%2Fb|c.lua",
                Command::Rename {
                    path: "a/b".into(),
                    name: "c.lua".into(),
                },
            ),
            ("refresh", Command::Refresh),
        ];
        for (line, want) in cases {
            assert_eq!(parse_command(line), Some(want), "{line}");
        }
        assert_eq!(parse_command(""), None);
        assert_eq!(parse_command("format-disk|c"), None);
        assert_eq!(percent_decode("%E2%9C%93%zz%4"), "\u{2713}%zz%4");
    }

    #[test]
    fn js_strings_cannot_escape_their_context() {
        let hostile = "a\"b\\c</script><img src=x onerror=alert(1)>&'\n\u{2028}\u{1F600}";
        let lit = js_string(hostile);
        assert!(lit.is_ascii());
        for bad in ["</", "<", ">", "&", "'", "\n", "\r"] {
            assert!(
                !lit[1..lit.len() - 1].contains(bad),
                "{bad:?} leaked into {lit}"
            );
        }
        assert!(
            lit.contains("\\ud83d\\ude00"),
            "astral characters become surrogate pairs: {lit}"
        );
        assert_eq!(js_string("plain/path.lua"), "\"plain/path.lua\"");
    }

    #[test]
    fn state_lists_entries_and_marks_scripts_that_run_at_load() {
        let b = bundle_of(&[("main.lua", "x"), ("lib/u.lua", "y")]);
        let view = crate::ui::view(&Status::Active {
            names: vec!["main.lua".into()],
        });
        let js = state_js(&PanelState {
            workbook: Some("Book \"1\".xlsx"),
            ok: true,
            hint: "",
            view: &view,
            dirty: true,
            fresh: false,
            message: Some((true, "boom <b>".into())),
            bundle: Some(&b),
        });
        assert!(js.starts_with("setState({") && js.ends_with("]});"));
        assert!(js.contains("light:\"yellow\"") && js.contains("dirty:true") && js.contains("n:2"));
        assert!(
            js.contains("{p:\"main.lua\",k:\"f\",s:1,r:1}")
                && js.contains("{p:\"lib\",k:\"d\",s:0,r:0}")
        );
        assert!(
            js.contains("{p:\"lib/u.lua\",k:\"f\",s:1,r:0}"),
            "only root scripts run at load: {js}"
        );
        assert!(js.contains("Book \\\"1\\\".xlsx") && !js.contains("<b>"));
    }

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
