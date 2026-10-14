//! Script storage: a custom XML part inside the .xlsx/.xlsm package.
//!
//! Layout written by [`embed_scripts`]:
//!
//! ```text
//! customXml/itemN.xml            <scripts xmlns="urn:xll-bridge:scripts:v1">…</scripts>
//! customXml/itemPropsN.xml       datastore item properties (required by OPC for custom XML)
//! customXml/_rels/itemN.xml.rels relationship item -> itemProps
//! xl/_rels/workbook.xml.rels     + relationship workbook -> item
//! [Content_Types].xml            + override for itemProps
//! ```
//!
//! Excel round-trips custom XML parts it doesn't understand, so scripts survive a save.
//! Reading never executes anything: it only parses XML.

use std::io::{Read, Seek, Write};
use std::path::Path;

use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use sha2::{Digest, Sha256};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

pub const SCRIPTS_NS: &str = "urn:xll-bridge:scripts:v1";

const CONTENT_TYPES: &str = "[Content_Types].xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const REL_TYPE_CUSTOM_XML: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const REL_TYPE_CUSTOM_XML_PROPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const CT_ITEM_PROPS: &str = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
/// Upper bound on any single part we read into memory.
const MAX_PART_BYTES: u64 = 16 * 1024 * 1024;

#[derive(Debug, thiserror::Error)]
pub enum PackageError {
    #[error("io: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("xml: {0}")]
    Xml(String),
    #[error("not a valid workbook package: {0}")]
    Invalid(String),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Script {
    pub name: String,
    pub code: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ScriptSet {
    pub scripts: Vec<Script>,
}

impl ScriptSet {
    pub fn is_empty(&self) -> bool {
        self.scripts.is_empty()
    }

    /// SHA-256 over names and code, in document order. This is what trust is pinned to,
    /// so any edit to any script re-triggers the banner.
    pub fn digest(&self) -> String {
        let mut h = Sha256::new();
        for s in &self.scripts {
            h.update((s.name.len() as u64).to_le_bytes());
            h.update(s.name.as_bytes());
            h.update((s.code.len() as u64).to_le_bytes());
            h.update(s.code.as_bytes());
        }
        h.finalize().iter().map(|b| format!("{b:02x}")).collect()
    }
}

// ---------------------------------------------------------------- XML (de)serialisation

fn escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            // XML 1.0 forbids most control characters; keep tab/LF, normalise CR away.
            '\r' => {}
            c if (c as u32) < 0x20 && c != '\t' && c != '\n' => {}
            c => out.push(c),
        }
    }
    out
}

pub fn scripts_to_xml(set: &ScriptSet) -> String {
    let mut out = String::from(XML_DECL);
    out.push_str(&format!("<scripts xmlns=\"{SCRIPTS_NS}\">\n"));
    for s in &set.scripts {
        out.push_str(&format!(
            "  <script name=\"{}\">{}</script>\n",
            escape(&s.name),
            escape(&s.code)
        ));
    }
    out.push_str("</scripts>\n");
    out
}

fn script_name(e: &quick_xml::events::BytesStart<'_>) -> Result<String, PackageError> {
    for a in e.attributes().flatten() {
        if a.key.as_ref() == "name" {
            return Ok(a
                .normalized_value(XmlVersion::Implicit1_0)
                .map_err(|e| PackageError::Xml(e.to_string()))?
                .into_owned());
        }
    }
    Ok(String::new())
}

/// Parses a custom XML item. `Ok(None)` means "valid XML, but not our part".
pub fn scripts_from_xml(xml: &str) -> Result<Option<ScriptSet>, PackageError> {
    let mut reader = Reader::from_str(xml);
    let mut set = ScriptSet::default();
    let mut seen_root = false;
    let mut current: Option<(String, String)> = None;

    loop {
        let event = reader
            .read_event()
            .map_err(|e| PackageError::Xml(e.to_string()))?;
        match event {
            Event::Start(e) => {
                let local = e.local_name();
                if !seen_root {
                    if local.as_ref() != "scripts" {
                        return Ok(None);
                    }
                    let ns_ok = e
                        .attributes()
                        .flatten()
                        .any(|a| a.key.as_ref() == "xmlns" && a.value.as_ref() == SCRIPTS_NS);
                    if !ns_ok {
                        return Ok(None);
                    }
                    seen_root = true;
                } else if local.as_ref() == "script" {
                    current = Some((script_name(&e)?, String::new()));
                }
            }
            Event::Empty(e) => {
                if !seen_root {
                    // `<scripts xmlns=".."/>`: our part, with nothing in it.
                    let ours = e.local_name().as_ref() == "scripts"
                        && e.attributes()
                            .flatten()
                            .any(|a| a.key.as_ref() == "xmlns" && a.value.as_ref() == SCRIPTS_NS);
                    return Ok(ours.then(ScriptSet::default));
                } else if e.local_name().as_ref() == "script" {
                    set.scripts.push(Script {
                        name: script_name(&e)?,
                        code: String::new(),
                    });
                }
            }
            Event::Text(t) => {
                if let Some((_, code)) = current.as_mut() {
                    let text = t.xml_content(XmlVersion::Implicit1_0);
                    code.push_str(&text);
                }
            }
            Event::GeneralRef(r) => {
                if let Some((_, code)) = current.as_mut() {
                    if let Ok(Some(c)) = r.resolve_char_ref() {
                        code.push(c);
                    } else {
                        match r.as_ref() {
                            "amp" => code.push('&'),
                            "lt" => code.push('<'),
                            "gt" => code.push('>'),
                            "quot" => code.push('"'),
                            "apos" => code.push('\''),
                            _ => {}
                        }
                    }
                }
            }
            Event::CData(c) => {
                if let Some((_, code)) = current.as_mut() {
                    code.push_str(&c.into_inner());
                }
            }
            Event::End(e) => {
                if e.local_name().as_ref() == "script"
                    && let Some((name, code)) = current.take()
                {
                    set.scripts.push(Script { name, code });
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(if seen_root { Some(set) } else { None })
}

// ---------------------------------------------------------------- package I/O

fn read_part<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, PackageError> {
    let file = match zip.by_name(name) {
        Ok(f) => f,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(e) => return Err(e.into()),
    };
    let mut buf = Vec::new();
    file.take(MAX_PART_BYTES).read_to_end(&mut buf)?;
    Ok(Some(buf))
}

fn is_item_part(name: &str) -> Option<u32> {
    let rest = name.strip_prefix("customXml/item")?.strip_suffix(".xml")?;
    rest.parse().ok() // "itemProps1" does not parse, which is what we want
}

/// Finds our scripts part in a workbook package. Returns the item number along with the scripts.
fn find_scripts<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
) -> Result<(Option<(u32, ScriptSet)>, u32), PackageError> {
    let names: Vec<String> = zip
        .file_names()
        .filter_map(|n| n.ok().map(|n| n.into_owned()))
        .collect();
    let mut max_item = 0;
    let mut found = None;
    for name in names {
        let Some(n) = is_item_part(&name) else {
            continue;
        };
        max_item = max_item.max(n);
        if found.is_some() {
            continue;
        }
        let Some(bytes) = read_part(zip, &name)? else {
            continue;
        };
        // Foreign custom XML parts may be anything; only ours matters, so ignore parse failures.
        if let Ok(xml) = std::str::from_utf8(&bytes)
            && let Ok(Some(set)) = scripts_from_xml(xml)
        {
            found = Some((n, set));
        }
    }
    Ok((found, max_item))
}

/// Reads the embedded scripts, if any. Does not run or compile anything.
pub fn read_scripts(path: &Path) -> Result<Option<ScriptSet>, PackageError> {
    let file = std::fs::File::open(path)?;
    let mut zip = ZipArchive::new(file)?;
    Ok(find_scripts(&mut zip)?.0.map(|(_, set)| set))
}

fn insert_before(xml: &str, closing_tag: &str, snippet: &str) -> Result<String, PackageError> {
    let idx = xml
        .rfind(closing_tag)
        .ok_or_else(|| PackageError::Invalid(format!("missing {closing_tag}")))?;
    Ok(format!("{}{}{}", &xml[..idx], snippet, &xml[idx..]))
}

/// Writes `dst` as a copy of `src` carrying `scripts`. An existing scripts part is replaced
/// in place; otherwise a new one is wired into the package. `src` and `dst` may not be equal.
pub fn embed_scripts(src: &Path, dst: &Path, scripts: &ScriptSet) -> Result<(), PackageError> {
    if src == dst {
        return Err(PackageError::Invalid(
            "source and destination must differ".into(),
        ));
    }
    let mut zip = ZipArchive::new(std::fs::File::open(src)?)?;
    let (existing, max_item) = find_scripts(&mut zip)?;
    let xml = scripts_to_xml(scripts);
    let opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    let mut out = ZipWriter::new(std::fs::File::create(dst)?);
    let result = (|| -> Result<(), PackageError> {
        match existing {
            Some((n, _)) => {
                let target = format!("customXml/item{n}.xml");
                for i in 0..zip.len() {
                    let f = zip.by_index_raw(i)?;
                    if f.name().is_ok_and(|n| n == target) {
                        continue;
                    }
                    out.raw_copy_file(f)?;
                }
                out.start_file(target, opts)?;
                out.write_all(xml.as_bytes())?;
            }
            None => {
                let n = max_item + 1;
                let ct = read_part(&mut zip, CONTENT_TYPES)?
                    .ok_or_else(|| PackageError::Invalid(format!("missing {CONTENT_TYPES}")))?;
                let rels = read_part(&mut zip, WORKBOOK_RELS)?
                    .ok_or_else(|| PackageError::Invalid(format!("missing {WORKBOOK_RELS}")))?;
                let ct = insert_before(
                    &String::from_utf8_lossy(&ct),
                    "</Types>",
                    &format!(
                        "<Override PartName=\"/customXml/itemProps{n}.xml\" ContentType=\"{CT_ITEM_PROPS}\"/>"
                    ),
                )?;
                let rels = insert_before(
                    &String::from_utf8_lossy(&rels),
                    "</Relationships>",
                    &format!(
                        "<Relationship Id=\"rIdXllBridge{n}\" Type=\"{REL_TYPE_CUSTOM_XML}\" Target=\"../customXml/item{n}.xml\"/>"
                    ),
                )?;
                for i in 0..zip.len() {
                    let f = zip.by_index_raw(i)?;
                    if f.name()
                        .is_ok_and(|n| n == CONTENT_TYPES || n == WORKBOOK_RELS)
                    {
                        continue;
                    }
                    out.raw_copy_file(f)?;
                }
                out.start_file(CONTENT_TYPES, opts)?;
                out.write_all(ct.as_bytes())?;
                out.start_file(WORKBOOK_RELS, opts)?;
                out.write_all(rels.as_bytes())?;
                out.start_file(format!("customXml/item{n}.xml"), opts)?;
                out.write_all(xml.as_bytes())?;
                out.start_file(format!("customXml/itemProps{n}.xml"), opts)?;
                out.write_all(
                    format!(
                        "{XML_DECL}<ds:datastoreItem ds:itemID=\"{{8F0B5C2E-4D1A-4E37-9B6A-5A7C1E3D2F10}}\" \
                         xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\">\
                         <ds:schemaRefs><ds:schemaRef ds:uri=\"{SCRIPTS_NS}\"/></ds:schemaRefs>\
                         </ds:datastoreItem>"
                    )
                    .as_bytes(),
                )?;
                out.start_file(format!("customXml/_rels/item{n}.xml.rels"), opts)?;
                out.write_all(
                    format!(
                        "{XML_DECL}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
                         <Relationship Id=\"rId1\" Type=\"{REL_TYPE_CUSTOM_XML_PROPS}\" Target=\"itemProps{n}.xml\"/>\
                         </Relationships>"
                    )
                    .as_bytes(),
                )?;
            }
        }
        Ok(())
    })();
    match result {
        Ok(()) => {
            out.finish()?;
            Ok(())
        }
        Err(e) => {
            drop(out);
            let _ = std::fs::remove_file(dst);
            Err(e)
        }
    }
}

#[cfg(test)]
pub(crate) mod testutil {
    use super::*;

    /// Smallest package Excel-style readers accept for our purposes.
    pub fn minimal_xlsx(path: &Path) {
        let mut z = ZipWriter::new(std::fs::File::create(path).unwrap());
        let o = SimpleFileOptions::default();
        z.start_file(CONTENT_TYPES, o).unwrap();
        z.write_all(
            br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#,
        )
        .unwrap();
        z.start_file(WORKBOOK_RELS, o).unwrap();
        z.write_all(
            br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
        )
        .unwrap();
        z.start_file("xl/workbook.xml", o).unwrap();
        z.write_all(b"<workbook/>").unwrap();
        z.finish().unwrap();
    }
}

#[cfg(test)]
mod tests {
    use super::testutil::minimal_xlsx;
    use super::*;

    fn sample() -> ScriptSet {
        ScriptSet {
            scripts: vec![
                Script {
                    name: "main".into(),
                    code: "function f(a) return a < 2 and \"x&y\" end".into(),
                },
                Script {
                    name: "util".into(),
                    code: "-- ]]> tricky\nx = 1\n".into(),
                },
            ],
        }
    }

    #[test]
    fn xml_round_trip_preserves_code() {
        let set = sample();
        let back = scripts_from_xml(&scripts_to_xml(&set)).unwrap().unwrap();
        assert_eq!(back.scripts[0].code, set.scripts[0].code);
        assert_eq!(back.scripts[1].code.replace('\r', ""), set.scripts[1].code);
        assert_eq!(back.digest(), set.digest());
    }

    #[test]
    fn foreign_xml_is_ignored() {
        assert_eq!(scripts_from_xml("<other xmlns=\"urn:x\"/>").unwrap(), None);
        assert_eq!(
            scripts_from_xml("<scripts xmlns=\"urn:other\"/>").unwrap(),
            None
        );
    }

    #[test]
    fn digest_changes_with_content() {
        let a = sample();
        let mut b = sample();
        b.scripts[0].code.push(' ');
        assert_ne!(a.digest(), b.digest());
    }

    #[test]
    fn embed_then_read_then_replace() {
        let dir = tempfile::tempdir().unwrap();
        let src = dir.path().join("in.xlsx");
        let one = dir.path().join("one.xlsx");
        let two = dir.path().join("two.xlsx");
        minimal_xlsx(&src);
        assert!(read_scripts(&src).unwrap().is_none());

        embed_scripts(&src, &one, &sample()).unwrap();
        assert_eq!(
            read_scripts(&one).unwrap().unwrap().digest(),
            sample().digest()
        );

        let replacement = ScriptSet {
            scripts: vec![Script {
                name: "n".into(),
                code: "x=2".into(),
            }],
        };
        embed_scripts(&one, &two, &replacement).unwrap();
        assert_eq!(read_scripts(&two).unwrap().unwrap(), replacement);

        // Replacing must not duplicate parts.
        let zip = ZipArchive::new(std::fs::File::open(&two).unwrap()).unwrap();
        let items = zip
            .file_names()
            .filter(|n| n.as_ref().is_ok_and(|n| is_item_part(n).is_some()))
            .count();
        assert_eq!(items, 1);
    }
}
