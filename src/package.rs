//! Embedded file storage: a custom XML part inside the .xlsx/.xlsm package.
//!
//! Layout written by [`embed_bundle`]:
//!
//! ```text
//! customXml/itemN.xml            <files xmlns="urn:xll-bridge:files:v1">…</files>
//! customXml/itemPropsN.xml       datastore item properties (required by OPC for custom XML)
//! customXml/_rels/itemN.xml.rels relationship item -> itemProps
//! xl/_rels/workbook.xml.rels     + relationship workbook -> item
//! [Content_Types].xml            + override for itemProps
//! ```
//!
//! Excel round-trips custom XML parts it doesn't understand, so the files survive a save, and the
//! same part is reachable at runtime through `Workbook.CustomXMLParts`. Reading never executes
//! anything: it only parses XML. The original flat `<scripts>` format is still read.

use std::io::{Read, Seek, Write};
use std::path::Path;

use base64::Engine;
use base64::engine::general_purpose::STANDARD as B64;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::vfs::{Bundle, is_xml_text};

pub const FILES_NS: &str = "urn:xll-bridge:files:v1";
/// The first, flat format: `<scripts><script name="x">code</script></scripts>`.
pub const LEGACY_NS: &str = "urn:xll-bridge:scripts:v1";

const CONTENT_TYPES: &str = "[Content_Types].xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const REL_TYPE_CUSTOM_XML: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const REL_TYPE_CUSTOM_XML_PROPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const CT_ITEM_PROPS: &str = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
/// Upper bound on any single part we read into memory (base64 inflates the 64 MB payload cap).
const MAX_PART_BYTES: u64 = 160 * 1024 * 1024;

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

fn xml_err(e: impl std::fmt::Display) -> PackageError {
    PackageError::Xml(e.to_string())
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
            c => out.push(c),
        }
    }
    out
}

pub fn bundle_to_xml(b: &Bundle) -> String {
    let mut out = String::from(XML_DECL);
    out.push_str(&format!("<files xmlns=\"{FILES_NS}\">\n"));
    let mut dirs: Vec<&String> = b.explicit_dirs().iter().collect();
    dirs.sort();
    for d in dirs {
        out.push_str(&format!("  <dir path=\"{}\"/>\n", escape(d)));
    }
    let mut files: Vec<_> = b.files().iter().collect();
    files.sort_by(|a, b| a.path.cmp(&b.path));
    for f in files {
        if is_xml_text(&f.data) {
            // `is_xml_text` guarantees valid UTF-8.
            let text = String::from_utf8_lossy(&f.data);
            out.push_str(&format!(
                "  <file path=\"{}\">{}</file>\n",
                escape(&f.path),
                escape(&text)
            ));
        } else {
            out.push_str(&format!(
                "  <file path=\"{}\" encoding=\"base64\">{}</file>\n",
                escape(&f.path),
                B64.encode(&f.data)
            ));
        }
    }
    out.push_str("</files>\n");
    out
}

fn attr(e: &BytesStart<'_>, name: &str) -> Result<Option<String>, PackageError> {
    for a in e.attributes().flatten() {
        if a.key.as_ref() == name {
            return Ok(Some(
                a.normalized_value(XmlVersion::Implicit1_0)
                    .map_err(xml_err)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn root_ns(e: &BytesStart<'_>) -> Option<String> {
    attr(e, "xmlns").ok().flatten()
}

enum Open {
    File {
        path: String,
        base64: bool,
        text: String,
    },
    Script {
        name: String,
        text: String,
    },
}

/// Parses a custom XML item. `Ok(None)` means "valid XML, but not ours".
pub fn bundle_from_xml(xml: &str) -> Result<Option<Bundle>, PackageError> {
    let mut reader = Reader::from_str(xml);
    let mut bundle = Bundle::new();
    let mut root: Option<&'static str> = None; // "files" | "scripts"
    let mut open: Option<Open> = None;
    let add_err = |e: crate::vfs::VfsError| PackageError::Invalid(e.to_string());

    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) if root.is_none() => {
                let ns = root_ns(&e);
                root = match (e.local_name().as_ref(), ns.as_deref()) {
                    ("files", Some(FILES_NS)) => Some("files"),
                    ("scripts", Some(LEGACY_NS)) => Some("scripts"),
                    _ => return Ok(None),
                };
            }
            Event::Start(e) => match (root, e.local_name().as_ref()) {
                (Some("files"), "file") => {
                    let path = attr(&e, "path")?.ok_or_else(|| xml_err("file without path"))?;
                    let base64 = attr(&e, "encoding")?.as_deref() == Some("base64");
                    open = Some(Open::File {
                        path,
                        base64,
                        text: String::new(),
                    });
                }
                (Some("scripts"), "script") => {
                    open = Some(Open::Script {
                        name: attr(&e, "name")?.unwrap_or_default(),
                        text: String::new(),
                    });
                }
                _ => {}
            },
            Event::Empty(e) => match (root, e.local_name().as_ref()) {
                (Some("files"), "dir") => {
                    let path = attr(&e, "path")?.ok_or_else(|| xml_err("dir without path"))?;
                    bundle.mkdir(&path).map_err(add_err)?;
                }
                (Some("files"), "file") => {
                    let path = attr(&e, "path")?.ok_or_else(|| xml_err("file without path"))?;
                    bundle.put_file(&path, Vec::new(), true).map_err(add_err)?;
                }
                (Some("scripts"), "script") => {
                    let name = attr(&e, "name")?.unwrap_or_default();
                    bundle
                        .put_file(&format!("{name}.lua"), Vec::new(), true)
                        .map_err(add_err)?;
                }
                _ => {}
            },
            Event::Text(t) => match open.as_mut() {
                Some(Open::File { text, .. } | Open::Script { text, .. }) => {
                    text.push_str(&t.xml_content(XmlVersion::Implicit1_0));
                }
                None => {}
            },
            Event::GeneralRef(r) => {
                if let Some(Open::File { text, .. } | Open::Script { text, .. }) = open.as_mut() {
                    if let Ok(Some(c)) = r.resolve_char_ref() {
                        text.push(c);
                    } else {
                        match r.as_ref() {
                            "amp" => text.push('&'),
                            "lt" => text.push('<'),
                            "gt" => text.push('>'),
                            "quot" => text.push('"'),
                            "apos" => text.push('\''),
                            _ => {}
                        }
                    }
                }
            }
            Event::CData(c) => {
                if let Some(Open::File { text, .. } | Open::Script { text, .. }) = open.as_mut() {
                    text.push_str(&c.into_inner());
                }
            }
            Event::End(e) => {
                let local = e.local_name();
                if matches!(local.as_ref(), "file" | "script") {
                    match open.take() {
                        Some(Open::File { path, base64, text }) => {
                            let data = if base64 {
                                let compact: String = text.split_whitespace().collect();
                                B64.decode(compact).map_err(xml_err)?
                            } else {
                                text.into_bytes()
                            };
                            bundle.put_file(&path, data, true).map_err(add_err)?;
                        }
                        Some(Open::Script { name, text }) => {
                            bundle
                                .put_file(&format!("{name}.lua"), text.into_bytes(), true)
                                .map_err(add_err)?;
                        }
                        None => {}
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(root.map(|_| bundle))
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

/// Finds our part in a workbook package: (item number, bundle) and the highest item number used.
fn find_bundle<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
) -> Result<(Option<(u32, Bundle)>, u32), PackageError> {
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
            && let Ok(Some(b)) = bundle_from_xml(xml)
        {
            found = Some((n, b));
        }
    }
    Ok((found, max_item))
}

/// Reads the embedded files from a saved workbook, if any. Does not run or compile anything.
pub fn read_bundle(path: &Path) -> Result<Option<Bundle>, PackageError> {
    let mut zip = ZipArchive::new(std::fs::File::open(path)?)?;
    Ok(find_bundle(&mut zip)?.0.map(|(_, b)| b))
}

fn insert_before(xml: &str, closing_tag: &str, snippet: &str) -> Result<String, PackageError> {
    let idx = xml
        .rfind(closing_tag)
        .ok_or_else(|| PackageError::Invalid(format!("missing {closing_tag}")))?;
    Ok(format!("{}{}{}", &xml[..idx], snippet, &xml[idx..]))
}

/// Writes `dst` as a copy of `src` carrying `bundle`. An existing part is replaced in place;
/// otherwise a new one is wired into the package. `src` and `dst` may not be equal.
pub fn embed_bundle(src: &Path, dst: &Path, bundle: &Bundle) -> Result<(), PackageError> {
    if src == dst {
        return Err(PackageError::Invalid(
            "source and destination must differ".into(),
        ));
    }
    let mut zip = ZipArchive::new(std::fs::File::open(src)?)?;
    let (existing, max_item) = find_bundle(&mut zip)?;
    let xml = bundle_to_xml(bundle);
    let opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    let mut out = ZipWriter::new(std::fs::File::create(dst)?);
    let result = (|| -> Result<(), PackageError> {
        match existing {
            Some((n, _)) => {
                let target = format!("customXml/item{n}.xml");
                for i in 0..zip.len() {
                    let f = zip.by_index_raw(i)?;
                    if f.name().is_ok_and(|name| name == target) {
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
                        .is_ok_and(|name| name == CONTENT_TYPES || name == WORKBOOK_RELS)
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
                         <ds:schemaRefs><ds:schemaRef ds:uri=\"{FILES_NS}\"/></ds:schemaRefs>\
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

    /// Convenience for tests: a bundle of (path, text) files.
    pub fn bundle_of(files: &[(&str, &str)]) -> Bundle {
        let mut b = Bundle::new();
        for (p, c) in files {
            b.put_file(p, c.as_bytes().to_vec(), true).unwrap();
        }
        b
    }
}

#[cfg(test)]
mod tests {
    use super::testutil::{bundle_of, minimal_xlsx};
    use super::*;

    fn sample() -> Bundle {
        let mut b = bundle_of(&[
            ("main.lua", "function f(a) return a < 2 and \"x&y\" end\n"),
            ("lib/util.lua", "-- ]]> tricky <b>\r\nx = 1\n"),
        ]);
        b.put_file("bin/native.dll", vec![0, 1, 2, 0xff, b'<', b'&', 13], true)
            .unwrap();
        b.mkdir("empty/nested").unwrap();
        b
    }

    #[test]
    fn xml_round_trip_is_exact_for_text_binary_and_folders() {
        let b = sample();
        let xml = bundle_to_xml(&b);
        assert!(
            xml.contains("encoding=\"base64\"") && xml.contains("<dir path=\"empty/nested\"/>")
        );
        let back = bundle_from_xml(&xml).unwrap().unwrap();
        assert_eq!(back, b);
        assert_eq!(back.digest(), b.digest());
        assert_eq!(
            back.get("bin/native.dll").unwrap().data,
            vec![0, 1, 2, 0xff, b'<', b'&', 13]
        );
    }

    #[test]
    fn xml_survives_a_reformatting_parser() {
        // Excel/MSXML may re-serialise: indentation between elements and attribute order change.
        let xml = format!(
            "<?xml version=\"1.0\"?>\n<files xmlns=\"{FILES_NS}\">\n\n<file encoding=\"base64\" path=\"a.bin\">\n AAEC\n /w==\n</file>\n<file path=\"t.txt\">hi &#13;there</file></files>"
        );
        // base64 is whitespace-tolerant; a literal CR via a character reference is kept as binary-ish text.
        let b = bundle_from_xml(&xml).unwrap().unwrap();
        assert_eq!(b.get("a.bin").unwrap().data, vec![0, 1, 2, 0xff]);
        assert_eq!(b.get("t.txt").unwrap().data, b"hi \rthere");
    }

    #[test]
    fn legacy_scripts_format_still_reads() {
        let xml = format!(
            "<scripts xmlns=\"{LEGACY_NS}\"><script name=\"main\">return 1</script><script name=\"e\"/></scripts>"
        );
        let b = bundle_from_xml(&xml).unwrap().unwrap();
        assert_eq!(b.get("main.lua").unwrap().data, b"return 1");
        assert!(b.get("e.lua").is_some());
    }

    #[test]
    fn foreign_and_hostile_xml() {
        assert_eq!(bundle_from_xml("<other xmlns=\"urn:x\"/>").unwrap(), None);
        assert_eq!(
            bundle_from_xml("<files xmlns=\"urn:other\"/>").unwrap(),
            None
        );
        let evil =
            format!("<files xmlns=\"{FILES_NS}\"><file path=\"../../evil.dll\">x</file></files>");
        assert!(
            bundle_from_xml(&evil).is_err(),
            "traversal paths must be rejected on read"
        );
        let bad_b64 = format!(
            "<files xmlns=\"{FILES_NS}\"><file path=\"a\" encoding=\"base64\">***</file></files>"
        );
        assert!(bundle_from_xml(&bad_b64).is_err());
    }

    #[test]
    fn embed_then_read_then_replace() {
        let dir = tempfile::tempdir().unwrap();
        let (src, one, two) = (
            dir.path().join("in.xlsx"),
            dir.path().join("one.xlsx"),
            dir.path().join("two.xlsx"),
        );
        minimal_xlsx(&src);
        assert!(read_bundle(&src).unwrap().is_none());

        embed_bundle(&src, &one, &sample()).unwrap();
        assert_eq!(read_bundle(&one).unwrap().unwrap(), sample());

        let replacement = bundle_of(&[("n.lua", "x=2")]);
        embed_bundle(&one, &two, &replacement).unwrap();
        assert_eq!(read_bundle(&two).unwrap().unwrap(), replacement);

        // Replacing must not duplicate parts.
        let zip = ZipArchive::new(std::fs::File::open(&two).unwrap()).unwrap();
        let items = zip
            .file_names()
            .filter(|n| n.as_ref().is_ok_and(|n| is_item_part(n).is_some()))
            .count();
        assert_eq!(items, 1);
    }
}
