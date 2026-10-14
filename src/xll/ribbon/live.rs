//! Reads and writes the embedded files of the *open* workbook through Excel's object model
//! (`Workbook.CustomXMLParts`), so Excel itself persists them on the next save.

use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;
use windows::core::{BSTR, Result};

use super::dispatch;
use crate::package::{FILES_NS, LEGACY_NS, bundle_from_xml, bundle_to_xml};
use crate::vfs::Bundle;

fn parts(wb: &IDispatch) -> Result<IDispatch> {
    IDispatch::try_from(&dispatch::get(wb, "CustomXMLParts")?)
}

/// Our parts in the workbook: (part object, parsed contents).
fn ours(wb: &IDispatch) -> Result<Vec<(IDispatch, Bundle)>> {
    let coll = parts(wb)?;
    let count = i32::try_from(&dispatch::get(&coll, "Count")?).unwrap_or(0);
    let mut found = Vec::new();
    for i in 1..=count {
        let Ok(part) = dispatch::get_with(&coll, "Item", &[VARIANT::from(i)])
            .and_then(|v| IDispatch::try_from(&v))
        else {
            continue;
        };
        let ns = dispatch::get(&part, "NamespaceURI")
            .ok()
            .and_then(|v| BSTR::try_from(&v).ok())
            .map(|b| b.to_string())
            .unwrap_or_default();
        if ns != FILES_NS && ns != LEGACY_NS {
            continue;
        }
        let xml = dispatch::get(&part, "XML")
            .ok()
            .and_then(|v| BSTR::try_from(&v).ok())
            .map(|b| b.to_string());
        if let Some(Ok(Some(bundle))) = xml.as_deref().map(bundle_from_xml) {
            found.push((part, bundle));
        }
    }
    Ok(found)
}

/// The workbook's current (possibly unsaved) embedded files; `None` if it has none.
pub fn read(wb: &IDispatch) -> Result<Option<Bundle>> {
    Ok(ours(wb)?.into_iter().next().map(|(_, b)| b))
}

/// Replaces the workbook's embedded files with `bundle` (removing the part when it is empty) and
/// marks the workbook as modified so Excel offers to save it.
pub fn write(wb: &IDispatch, bundle: &Bundle) -> Result<()> {
    let existing = ours(wb)?;
    // Add before deleting so a failed Add (e.g. a protected workbook) loses nothing.
    if bundle.is_empty() {
        for (part, _) in existing {
            dispatch::call(&part, "Delete", &[])?;
        }
    } else {
        let added = dispatch::call(
            &parts(wb)?,
            "Add",
            &[VARIANT::from(bundle_to_xml(bundle).as_str())],
        )?;
        drop(added);
        for (part, _) in existing {
            dispatch::call(&part, "Delete", &[])?;
        }
    }
    dispatch::put(wb, "Saved", VARIANT::from(false))
}
