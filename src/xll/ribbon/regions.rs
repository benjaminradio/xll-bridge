//! Formula regions in Excel: the object-model side of [`crate::regions`] and the ribbon actions.
//!
//! A region is a hidden workbook name (`_xs_<label>`) that Excel keeps pointing at the right
//! cells through inserts and deletes; the formula itself is the top-left (anchor) cell's own.

use std::path::Path;

use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;

use super::cells::{
    com, count_of, describe, dispatch_of, open_book, read_variant, resolve, string_of,
    write_variant,
};
use super::dispatch;
use crate::regions::{
    self, Axis, CheckReport, DevKind, OldCells, Rect, RegionInfo, RollMode, Rule, default_axis,
    label_of, name_for,
};
use crate::value::Value;
use crate::xlapi::{RegionOp, Reply, XlError, describe_deviations};

pub(super) struct Region {
    pub label: String,
    pub name: IDispatch,
    pub range: IDispatch,
    pub info: RegionInfo,
}

fn book_of(range: &IDispatch) -> String {
    dispatch_of(dispatch::get(range, "Worksheet"))
        .and_then(|ws| dispatch_of(dispatch::get(&ws, "Parent")))
        .and_then(|wb| string_of(dispatch::get(&wb, "Name")))
        .unwrap_or_default()
}

/// The regions of a workbook, found through its hidden names.
pub(super) fn all(wb: &IDispatch) -> Result<Vec<Region>, XlError> {
    let names = dispatch_of(dispatch::get(wb, "Names"))?;
    let n = count_of(&names, "Count")?;
    let mut out = Vec::new();
    for i in 1..=n {
        let Ok(nm) = dispatch_of(dispatch::get_with(&names, "Item", &[VARIANT::from(i)])) else {
            continue;
        };
        let full = string_of(dispatch::get(&nm, "Name")).unwrap_or_default();
        let Some(label) = label_of(&full) else {
            continue;
        };
        // A name whose cells were deleted points nowhere: skip it (Remove can still clear it).
        let Ok(range) = dispatch_of(dispatch::get(&nm, "RefersToRange")) else {
            continue;
        };
        let Ok(info) = describe(&range, &book_of(&range), false) else {
            continue;
        };
        let rule = rule_of(&nm);
        out.push(Region {
            label: label.to_owned(),
            name: nm,
            range,
            info: RegionInfo {
                name: label.to_owned(),
                address: info.address,
                rows: info.rows,
                cols: info.cols,
                rule,
            },
        });
    }
    Ok(out)
}

/// The roll rule kept in the comment of a region's hidden name.
fn rule_of(name: &IDispatch) -> Rule {
    string_of(dispatch::get(name, "Comment"))
        .map(|c| Rule::parse(&c))
        .unwrap_or_default()
}

pub(super) fn set_rule(r: &Region, rule: Rule) -> Result<(), XlError> {
    dispatch::put(&r.name, "Comment", VARIANT::from(rule.encode().as_str()))
        .map(drop)
        .map_err(com)
}

fn find(wb: &IDispatch, label: &str) -> Result<Region, XlError> {
    all(wb)?
        .into_iter()
        .find(|r| r.label == label)
        .ok_or_else(|| XlError::new(format!("no formula region named '{label}'")))
}

fn cell_address(range: &IDispatch, row: usize, col: usize) -> String {
    dispatch_of(dispatch::get_with(
        range,
        "Cells",
        &[VARIANT::from(row as i32 + 1), VARIANT::from(col as i32 + 1)],
    ))
    .and_then(|c| {
        string_of(dispatch::get_with(
            &c,
            "Address",
            &[VARIANT::from(false), VARIANT::from(false)],
        ))
    })
    .unwrap_or_else(|_| format!("R{}C{}", row + 1, col + 1))
}

/// The R1C1 formulas (or constants) of a block, row-major.
fn grid(range: &IDispatch) -> Result<Vec<Value>, XlError> {
    let v = dispatch::get(range, "FormulaR1C1").map_err(com)?;
    Ok(match read_variant(&v)? {
        Value::Array(a) => a.data,
        scalar => vec![scalar],
    })
}

pub(super) fn check(r: &Region) -> Result<CheckReport, XlError> {
    let cells = grid(&r.range)?;
    let anchor = match cells.first().map(regions::classify) {
        Some((DevKind::Formula, f)) => f,
        _ => {
            return Err(XlError::new(format!(
                "the anchor (top-left cell) of region '{}' holds no formula",
                r.label
            )));
        }
    };
    let deviations = regions::compare(&anchor, &cells, r.info.cols as usize, |i, j| {
        cell_address(&r.range, i, j)
    });
    Ok(CheckReport {
        region: r.info.clone(),
        anchor,
        deviations,
    })
}

fn fill(range: &IDispatch, formula: &str) -> Result<(), XlError> {
    dispatch::put(range, "FormulaR1C1", VARIANT::from(formula)).map_err(com)
}

/// Fills the region from its anchor; refuses (listing the cells) when something else is in the way
/// and `overwrite` is false.
pub(super) fn apply(r: &Region, overwrite: bool) -> Result<CheckReport, XlError> {
    let report = check(r)?;
    if report.non_blank() > 0 && !overwrite {
        return Err(XlError::new(describe_deviations(&report)));
    }
    if !report.deviations.is_empty() {
        fill(&r.range, &report.anchor)?;
    }
    Ok(report)
}

fn range_address(range: &IDispatch) -> Result<String, XlError> {
    let book = book_of(range);
    Ok(describe(range, &book, false)?.address)
}

fn overlaps(app: &IDispatch, a: &IDispatch, b: &IDispatch) -> bool {
    dispatch::call(
        app,
        "Intersect",
        &[VARIANT::from(a.clone()), VARIANT::from(b.clone())],
    )
    .ok()
    .and_then(|v| IDispatch::try_from(&v).ok())
    .is_some()
}

fn application(obj: &IDispatch) -> Result<IDispatch, XlError> {
    dispatch_of(dispatch::get(obj, "Application"))
}

fn default_label(range: &IDispatch, info_address: &str) -> String {
    let sheet = info_address
        .split('!')
        .next()
        .unwrap_or("Region")
        .trim_matches('\'');
    format!("{sheet}_{}", cell_address(range, 0, 0))
}

/// Registers `range` as a region (its top-left cell is the anchor) without filling it yet.
fn create(wb: &IDispatch, range: &IDispatch, label: Option<String>) -> Result<Region, XlError> {
    let info = describe(range, &book_of(range), false)?;
    if u64::from(info.rows) * u64::from(info.cols) < 2 {
        return Err(XlError::new("a formula region needs at least two cells"));
    }
    let label = label.unwrap_or_else(|| default_label(range, &info.address));
    let name = name_for(&label).map_err(XlError::new)?;
    let label = label_of(&name).unwrap_or_default().to_owned();
    let app = application(wb)?;
    for other in all(wb)? {
        if other.label == label {
            return Err(XlError::new(format!(
                "a formula region named '{label}' exists already"
            )));
        }
        if overlaps(&app, range, &other.range) {
            return Err(XlError::new(format!(
                "the range overlaps the formula region '{}'. {OVERLAP_HINT}",
                other.label
            )));
        }
    }
    let names = dispatch_of(dispatch::get(wb, "Names"))?;
    let nm = dispatch_of(dispatch::call(
        &names,
        "Add",
        &[
            VARIANT::from(name.as_str()),
            VARIANT::from(format!("={}", info.address).as_str()),
        ],
    ))?;
    dispatch::put(&nm, "Visible", VARIANT::from(false)).map_err(com)?;
    Ok(Region {
        info: RegionInfo {
            name: label.clone(),
            address: info.address,
            rows: info.rows,
            cols: info.cols,
            rule: Rule::default(),
        },
        label,
        name: nm,
        range: range.clone(),
    })
}

const OVERLAP_HINT: &str = "To redefine a region, select a block that starts at its anchor (top-left cell). To move the anchor, Remove the region, move the formula, then Define again";

/// Makes `range` a region and fills it from its anchor; on failure nothing is left behind.
///
/// A range that starts at the anchor of an existing region **redefines** that region (like
/// removing it and defining the range, but keeping its name and roll rule); the anchor never
/// moves. Any other range that overlaps a region is refused. The flag says whether it was a
/// redefinition.
pub(super) fn define(
    wb: &IDispatch,
    range: &IDispatch,
    label: Option<String>,
    overwrite: bool,
) -> Result<(Region, bool), XlError> {
    let first = |r: &IDispatch| -> Result<String, XlError> {
        let c = dispatch_of(dispatch::get_with(
            r,
            "Cells",
            &[VARIANT::from(1i32), VARIANT::from(1i32)],
        ))?;
        range_address(&c)
    };
    let top = first(range)?;
    let existing = all(wb)?
        .into_iter()
        .find(|r| first(&r.range).is_ok_and(|a| a == top));
    if let Some(old) = existing {
        return redefine(wb, old, range, overwrite).map(|r| (r, true));
    }
    let region = create(wb, range, label)?;
    if let Err(e) = apply(&region, overwrite) {
        let _ = dispatch::call(&region.name, "Delete", &[]);
        return Err(e);
    }
    Ok((region, false))
}

/// Points an existing region at a new range with the same anchor and fills it.
fn redefine(
    wb: &IDispatch,
    old: Region,
    range: &IDispatch,
    overwrite: bool,
) -> Result<Region, XlError> {
    let info = describe(range, &book_of(range), false)?;
    if u64::from(info.rows) * u64::from(info.cols) < 2 {
        return Err(XlError::new("a formula region needs at least two cells"));
    }
    let app = application(wb)?;
    for other in all(wb)? {
        if other.label != old.label && overlaps(&app, range, &other.range) {
            return Err(XlError::new(format!(
                "the range overlaps the formula region '{}'. {OVERLAP_HINT}",
                other.label
            )));
        }
    }
    let previous = string_of(dispatch::get(&old.name, "RefersTo")).ok();
    dispatch::put(
        &old.name,
        "RefersTo",
        VARIANT::from(format!("={}", info.address).as_str()),
    )
    .map_err(com)?;
    let region = Region {
        label: old.label.clone(),
        name: old.name.clone(),
        range: range.clone(),
        info: RegionInfo {
            name: old.label.clone(),
            address: info.address,
            rows: info.rows,
            cols: info.cols,
            rule: old.info.rule,
        },
    };
    if let Err(e) = apply(&region, overwrite) {
        if let Some(p) = previous {
            let _ = dispatch::put(&old.name, "RefersTo", VARIANT::from(p.as_str()));
        }
        return Err(e);
    }
    Ok(region)
}

fn rect_range(top_left: &IDispatch, r: Rect) -> Result<IDispatch, XlError> {
    let moved = dispatch_of(dispatch::get_with(
        top_left,
        "Offset",
        &[VARIANT::from(r.row), VARIANT::from(r.col)],
    ))
    .map_err(|_| XlError::new("the region would run off the sheet"))?;
    dispatch_of(dispatch::get_with(
        &moved,
        "Resize",
        &[VARIANT::from(r.rows), VARIANT::from(r.cols)],
    ))
}

/// Moves the start and/or end of the region. New cells get the anchor's formula; cells that
/// leave can be frozen to values.
pub(super) fn roll(
    r: &Region,
    start: i32,
    end: i32,
    old: OldCells,
    overwrite: bool,
) -> Result<Region, XlError> {
    let (rows, cols) = (r.info.rows as i32, r.info.cols as i32);
    let axis: Axis = default_axis(rows, cols);
    let plan = regions::roll(rows, cols, axis, start, end).map_err(XlError::new)?;
    let report = check(r)?;
    let anchor = report.anchor;
    let top_left = dispatch_of(dispatch::get_with(
        &r.range,
        "Cells",
        &[VARIANT::from(1i32), VARIANT::from(1i32)],
    ))?;
    // Look before touching anything: the cells the region grows into.
    let entering: Vec<IDispatch> = plan
        .entering
        .iter()
        .map(|rect| rect_range(&top_left, *rect))
        .collect::<Result<_, _>>()?;
    let new_range = rect_range(&top_left, plan.new)?;
    let mut in_the_way = Vec::new();
    for rng in &entering {
        let cells = grid(rng)?;
        let cols = describe(rng, &book_of(rng), false)?.cols as usize;
        in_the_way.extend(regions::compare(&anchor, &cells, cols, |i, j| {
            cell_address(rng, i, j)
        }));
    }
    let blocking = in_the_way
        .iter()
        .filter(|d| d.kind != DevKind::Blank)
        .count();
    if blocking > 0 && !overwrite {
        let probe = CheckReport {
            region: r.info.clone(),
            anchor: anchor.clone(),
            deviations: in_the_way,
        };
        return Err(XlError::new(describe_deviations(&probe)));
    }
    for rect in &plan.leaving {
        let rng = rect_range(&top_left, *rect)?;
        match old {
            OldCells::Keep => {}
            OldCells::Values => {
                let v = dispatch::get(&rng, "Value2").map_err(com)?;
                dispatch::put(&rng, "Value2", v).map_err(com)?;
            }
            OldCells::Clear => {
                dispatch::call(&rng, "ClearContents", &[]).map_err(com)?;
            }
        }
    }
    for rng in &entering {
        fill(rng, &anchor)?;
    }
    let address = range_address(&new_range)?;
    dispatch::put(
        &r.name,
        "RefersTo",
        VARIANT::from(format!("={address}").as_str()),
    )
    .map_err(com)?;
    let info = describe(&new_range, &book_of(&new_range), false)?;
    Ok(Region {
        label: r.label.clone(),
        name: r.name.clone(),
        range: new_range,
        info: RegionInfo {
            name: r.label.clone(),
            address: info.address,
            rows: info.rows,
            cols: info.cols,
            rule: r.info.rule,
        },
    })
}

pub(super) fn remove(r: &Region) -> Result<(), XlError> {
    dispatch::call(&r.name, "Delete", &[])
        .map(drop)
        .map_err(com)
}

/// The region that contains `cell`, if any.
pub(super) fn at(wb: &IDispatch, cell: &IDispatch) -> Result<Option<Region>, XlError> {
    let app = application(wb)?;
    Ok(all(wb)?
        .into_iter()
        .find(|r| overlaps(&app, cell, &r.range)))
}

// ------------------------------------------------------------------ the Lua / XlApi side

pub(super) fn run(app: &IDispatch, workbook: &Path, op: RegionOp) -> Result<Reply, XlError> {
    let book = super::cells::book_name_of(workbook)
        .ok_or_else(|| XlError::new("this workbook is not open in Excel"))?;
    let wb = open_book(app, &book)?;
    let one = |r: &Region| Reply::Regions(vec![r.info.clone()]);
    match op {
        RegionOp::List => Ok(Reply::Regions(
            all(&wb)?.into_iter().map(|r| r.info).collect(),
        )),
        RegionOp::At { address } => {
            let cell = resolve(app, workbook, &address)?.range;
            Ok(Reply::Regions(
                at(&wb, &cell)?.map(|r| r.info).into_iter().collect(),
            ))
        }
        RegionOp::Check { name } => Ok(Reply::Report(check(&find(&wb, &name)?)?)),
        RegionOp::Apply { name, overwrite } => {
            let r = find(&wb, &name)?;
            apply(&r, overwrite)?;
            Ok(one(&r))
        }
        RegionOp::Define {
            address,
            label,
            overwrite,
        } => {
            let range = resolve(app, workbook, &address)?.range;
            Ok(one(&define(&wb, &range, label, overwrite)?.0))
        }
        RegionOp::Roll {
            name,
            start,
            end,
            old,
            overwrite,
        } => {
            let r = find(&wb, &name)?;
            Ok(one(&roll(&r, start, end, old, overwrite)?))
        }
        RegionOp::Remove { name } => {
            remove(&find(&wb, &name)?)?;
            Ok(Reply::Done)
        }
        RegionOp::SetRule { name, rule } => {
            let mut r = find(&wb, &name)?;
            set_rule(&r, rule)?;
            r.info.rule = rule;
            Ok(one(&r))
        }
    }
}

// ------------------------------------------------------------------ ribbon actions

use windows::Win32::UI::WindowsAndMessaging::{
    MB_DEFBUTTON2, MB_ICONINFORMATION, MB_ICONWARNING, MB_OK, MB_YESNO,
};

use super::folder::{Choice, ask3};
use super::{excel_hwnd, message, set_status_bar};

const BLOCKED_MARK: &str = "hold something other than its formula";

fn is_blocked(e: &XlError) -> bool {
    e.message.contains(BLOCKED_MARK)
}

fn active_workbook(app: &IDispatch) -> Result<IDispatch, String> {
    dispatch_of(dispatch::get(app, "ActiveWorkbook"))
        .map_err(|_| "Open a workbook first.".to_owned())
}

fn region_here(app: &IDispatch, wb: &IDispatch) -> Result<Region, String> {
    let cell = dispatch_of(dispatch::get(app, "ActiveCell"))
        .map_err(|_| "Select a cell inside a formula region first.".to_owned())?;
    at(wb, &cell).map_err(|e| e.message)?.ok_or_else(|| {
        "The active cell is not in a formula region. Select a cell inside one.".to_owned()
    })
}

/// Runs `f`; if cells that hold something else are in the way and the setting says to ask, shows
/// them and runs `f` again with overwriting if the user agrees. `Ok(None)` = the user said no.
fn with_confirm<T>(
    app: &IDispatch,
    f: impl Fn(bool) -> Result<T, XlError>,
) -> Result<Option<T>, XlError> {
    let overwrite = crate::settings::load().overwrite_regions;
    match f(overwrite) {
        Ok(v) => Ok(Some(v)),
        Err(e) if is_blocked(&e) && !overwrite => {
            let text = e
                .message
                .split("; pass overwrite")
                .next()
                .unwrap_or(&e.message);
            if message(
                excel_hwnd(app),
                &format!("{text}.\n\nReplace them with the region's formula?"),
                "Formula region",
                MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2,
            ) {
                f(true).map(Some)
            } else {
                Ok(None)
            }
        }
        Err(e) => Err(e),
    }
}

fn fail(app: &IDispatch, what: &str, why: &str) {
    message(
        excel_hwnd(app),
        &format!("{what}\n\n{why}"),
        "Formula region",
        MB_OK | MB_ICONINFORMATION,
    );
}

/// "Define": the selection becomes a region, or redefines the region whose anchor it starts at.
pub(super) fn ui_define(app: &IDispatch) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        let sel = dispatch_of(dispatch::get(app, "Selection")).map_err(|_| {
            "Select the cells of the region first (the formula goes in its top-left cell)."
                .to_owned()
        })?;
        match with_confirm(app, |o| define(&wb, &sel, None, o)).map_err(|e| e.message)? {
            Some((r, again)) => Ok(format!(
                "Formula region '{}' {}: {} ({} cells)",
                r.label,
                if again { "redefined" } else { "defined" },
                r.info.address,
                u64::from(r.info.rows) * u64::from(r.info.cols)
            )),
            None => Ok("Formula region not defined".to_owned()),
        }
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not define the formula region.", &e),
    }
}

/// "Apply": fill the region around the active cell from its anchor again.
pub(super) fn ui_apply(app: &IDispatch) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        let r = region_here(app, &wb)?;
        match with_confirm(app, |o| apply(&r, o)).map_err(|e| e.message)? {
            Some(rep) if rep.deviations.is_empty() => Ok(format!(
                "Formula region '{}' is already consistent",
                r.label
            )),
            Some(rep) => Ok(format!(
                "Formula region '{}': {} cell(s) set to {}",
                r.label,
                rep.deviations.len(),
                rep.anchor
            )),
            None => Ok("Nothing changed".to_owned()),
        }
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not apply the formula region.", &e),
    }
}

/// "Remove": forgets the regions whose anchor lies in the selection (so a sloppy selection over
/// several is enough); with none, the region around the active cell. The cells keep their formulas.
pub(super) fn ui_remove(app: &IDispatch) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        let sel = dispatch_of(dispatch::get(app, "Selection"))
            .map_err(|_| "Select cells first.".to_owned())?;
        let mut hit: Vec<Region> = Vec::new();
        for r in all(&wb).map_err(|e| e.message)? {
            let anchor = dispatch_of(dispatch::get_with(
                &r.range,
                "Cells",
                &[VARIANT::from(1i32), VARIANT::from(1i32)],
            ))
            .map_err(|e| e.message)?;
            if overlaps(app, &sel, &anchor) {
                hit.push(r);
            }
        }
        if hit.is_empty() {
            hit.push(region_here(app, &wb)?);
        }
        for r in &hit {
            remove(r).map_err(|e| e.message)?;
        }
        Ok(match hit.as_slice() {
            [r] => format!(
                "Formula region '{}' removed; its cells keep their formulas",
                r.label
            ),
            _ => format!(
                "{} removed; their cells keep their formulas",
                plural(hit.len(), "formula region", "formula regions")
            ),
        })
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not remove the formula regions.", &e),
    }
}

fn select_cells(app: &IDispatch, r: &Region, addresses: &[String]) -> Result<(), XlError> {
    let sheet = dispatch_of(dispatch::get(&r.range, "Worksheet"))?;
    let mut union: Option<IDispatch> = None;
    for a in addresses.iter().take(200) {
        let cell = dispatch_of(dispatch::get_with(
            &sheet,
            "Range",
            &[VARIANT::from(a.as_str())],
        ))?;
        union = Some(match union {
            None => cell,
            Some(u) => dispatch_of(dispatch::call(
                app,
                "Union",
                &[VARIANT::from(u), VARIANT::from(cell)],
            ))?,
        });
    }
    if let Some(u) = union {
        dispatch::call(&sheet, "Activate", &[]).map_err(com)?;
        dispatch::call(&u, "Select", &[]).map_err(com)?;
    }
    Ok(())
}

/// Checks `regions`; lists what differs and, when asked, selects the first region's cells.
fn check_and_report(app: &IDispatch, regions: &[Region], highlight: bool) -> Result<(), String> {
    let mut reports = Vec::new();
    for r in regions {
        match check(r) {
            Ok(rep) => reports.push((r, Ok(rep))),
            Err(e) => reports.push((r, Err(e.message))),
        }
    }
    if highlight {
        clear_marks(app, MARK_DIFF).map_err(|e| e.message)?;
        for (r, rep) in &reports {
            if let Ok(rep) = rep
                && !rep.deviations.is_empty()
            {
                let addrs: Vec<String> = rep.deviations.iter().map(|d| d.address.clone()).collect();
                mark_cells(app, r, &addrs, MARK_DIFF, COLOR_DIFF).map_err(|e| e.message)?;
            }
        }
    }
    let bad: Vec<_> = reports
        .iter()
        .filter(|(_, rep)| rep.as_ref().map_or(true, |x| !x.deviations.is_empty()))
        .collect();
    if bad.is_empty() {
        set_status_bar(
            app,
            &format!("All {} formula region(s) are consistent", regions.len()),
        );
        return Ok(());
    }
    let mut text = format!(
        "{} of {} formula region(s) have cells that differ from their anchor:\n",
        bad.len(),
        regions.len()
    );
    for (r, rep) in bad.iter().take(6) {
        match rep {
            Err(e) => text.push_str(&format!("\n{} ({}): {e}", r.label, r.info.address)),
            Ok(rep) => {
                let list: Vec<String> = rep
                    .deviations
                    .iter()
                    .take(5)
                    .map(|d| format!("{} ({})", d.address, d.kind.word()))
                    .collect();
                let more = rep.deviations.len().saturating_sub(5);
                text.push_str(&format!(
                    "\n\u{2022} {} ({}): {}{}",
                    r.label,
                    r.info.address,
                    list.join(", "),
                    if more > 0 {
                        format!(" and {more} more")
                    } else {
                        String::new()
                    }
                ));
            }
        }
    }
    if highlight {
        text.push_str("\n\nThe differing cells are highlighted (Show \u{25B8} Clear removes it).");
        message(
            excel_hwnd(app),
            &text,
            "Formula regions",
            MB_OK | MB_ICONWARNING,
        );
        return Ok(());
    }
    text.push_str("\n\nSelect the cells of the first region that differ?");
    if message(
        excel_hwnd(app),
        &text,
        "Formula regions",
        MB_YESNO | MB_ICONWARNING,
    ) && let Some((r, Ok(rep))) = bad.iter().find(|(_, rep)| rep.is_ok())
    {
        let addrs: Vec<String> = rep.deviations.iter().map(|d| d.address.clone()).collect();
        select_cells(app, r, &addrs).map_err(|e| e.message)?;
    }
    Ok(())
}

const NO_REGIONS: &str = "There are no formula regions in this workbook. Type the formula in the top-left cell of a block, select the block and use Define.";

/// "Check All" (`all`), "Check" (the region around the active cell) and "Highlight differences"
/// (`diff`: Check All, marking the cells that differ). `tag` is `all`, `one` or `diff`.
pub(super) fn ui_check(app: &IDispatch, tag: &str) {
    let go = || -> Result<(), String> {
        let wb = active_workbook(app)?;
        let regions = if tag == "one" {
            vec![region_here(app, &wb)?]
        } else {
            all(&wb).map_err(|e| e.message)?
        };
        if regions.is_empty() {
            return Err(NO_REGIONS.into());
        }
        check_and_report(app, &regions, tag == "diff")
    };
    if let Err(e) = go() {
        fail(app, "Formula regions", &e);
    }
}

/// "Apply All": every region of the workbook is filled from its anchor again. Regions that hold
/// something else are listed and filled only if the user agrees (or the setting says to).
pub(super) fn ui_apply_all(app: &IDispatch) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        let regs = all(&wb).map_err(|e| e.message)?;
        if regs.is_empty() {
            return Err(NO_REGIONS.into());
        }
        let overwrite = crate::settings::load().overwrite_regions;
        let _ = dispatch::put(app, "ScreenUpdating", VARIANT::from(false));
        let (mut cells, mut regions_done) = (0usize, 0usize);
        let mut blocked: Vec<(Region, CheckReport)> = Vec::new();
        let mut failed: Vec<String> = Vec::new();
        for r in regs {
            match check(&r) {
                Ok(rep) if rep.deviations.is_empty() => {}
                Ok(rep) if rep.non_blank() > 0 && !overwrite => blocked.push((r, rep)),
                Ok(_) => match apply(&r, true) {
                    Ok(rep) => {
                        cells += rep.deviations.len();
                        regions_done += 1;
                    }
                    Err(e) => failed.push(format!("{}: {}", r.label, e.message)),
                },
                Err(e) => failed.push(format!("{}: {}", r.label, e.message)),
            }
        }
        if !blocked.is_empty() {
            let held: usize = blocked.iter().map(|(_, rep)| rep.non_blank()).sum();
            let names: Vec<String> = blocked
                .iter()
                .take(12)
                .map(|(r, rep)| format!("{} ({} cells)", r.label, rep.non_blank()))
                .collect();
            let text = format!(
                "{} hold {held} cell(s) other than their formula (hardcoded values or different formulas):\n\n  {}{}\n\nReplace them with the regions' formulas?",
                plural(blocked.len(), "region", "regions"),
                names.join("\n  "),
                if blocked.len() > 12 {
                    format!("\n  ... and {} more", blocked.len() - 12)
                } else {
                    String::new()
                }
            );
            if message(
                excel_hwnd(app),
                &text,
                "Apply all formula regions",
                MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2,
            ) {
                for (r, _) in blocked.drain(..) {
                    match apply(&r, true) {
                        Ok(rep) => {
                            cells += rep.deviations.len();
                            regions_done += 1;
                        }
                        Err(e) => failed.push(format!("{}: {}", r.label, e.message)),
                    }
                }
            }
        }
        let _ = dispatch::put(app, "ScreenUpdating", VARIANT::from(true));
        if !failed.is_empty() {
            let shown: Vec<&str> = failed.iter().take(10).map(String::as_str).collect();
            fail(
                app,
                &format!(
                    "{} could not be applied:",
                    plural(failed.len(), "region", "regions")
                ),
                &shown.join("\n"),
            );
        }
        let mut t = format!(
            "Applied: {} set in {}",
            plural(cells, "cell", "cells"),
            plural(regions_done, "region", "regions")
        );
        if !blocked.is_empty() {
            t += &format!("; {} skipped", blocked.len());
        }
        Ok(t)
    };
    let result = go();
    let _ = dispatch::put(app, "ScreenUpdating", VARIANT::from(true));
    match result {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not apply the formula regions.", &e),
    }
}

/// "Remove All": forgets every formula region on the active sheet (the cells keep their formulas).
pub(super) fn ui_remove_all(app: &IDispatch) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        let sheet = dispatch_of(dispatch::get(app, "ActiveSheet"))
            .map_err(|_| "Open a worksheet first.".to_owned())?;
        let sheet_name = string_of(dispatch::get(&sheet, "Name")).map_err(|e| e.message)?;
        let here: Vec<Region> = all(&wb)
            .map_err(|e| e.message)?
            .into_iter()
            .filter(|r| {
                dispatch_of(dispatch::get(&r.range, "Worksheet"))
                    .and_then(|w| string_of(dispatch::get(&w, "Name")))
                    .is_ok_and(|n| n == sheet_name)
            })
            .collect();
        if here.is_empty() {
            return Err(format!("There are no formula regions on '{sheet_name}'."));
        }
        if !message(
            excel_hwnd(app),
            &format!(
                "Remove {} on '{sheet_name}'?\n\nThe cells keep their formulas; only the regions (and their roll rules) are forgotten.",
                plural(here.len(), "formula region", "formula regions")
            ),
            "Remove all formula regions",
            MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2,
        ) {
            return Ok("Nothing changed".into());
        }
        for r in &here {
            remove(r).map_err(|e| e.message)?;
        }
        Ok(format!(
            "Removed {} on '{sheet_name}'",
            plural(here.len(), "formula region", "formula regions")
        ))
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not remove the formula regions.", &e),
    }
}

// ------------------------------------------------------------------ marking regions on the sheet
//
// Highlights and outlines are conditional formats, not cell formatting: they leave a cell's own
// formatting alone, follow the cells when rows or columns are inserted, and are recognised (and
// removed) by a tag in their formula.

const MARK_ANCHOR: &str = "xs-anchor";
const MARK_OUTLINE: &str = "xs-outline";
const MARK_DIFF: &str = "xs-diff";
const COLOR_ANCHOR: i32 = 0x9CEBFF; // light amber (BGR)
const COLOR_DIFF: i32 = 0x9CC7FF; // light orange (BGR)
const COLOR_OUTLINE: i32 = 0xC07000; // blue (BGR)

/// Adds a conditional format over `target` whose formula is `formula` (with a tag so it can be
/// found again); returns it so the caller can style it.
fn add_condition(target: &IDispatch, formula: &str) -> Result<IDispatch, XlError> {
    let fcs = dispatch_of(dispatch::get(target, "FormatConditions"))?;
    let fc = dispatch_of(dispatch::call(
        &fcs,
        "Add",
        &[
            VARIANT::from(2i32), // xlExpression
            super::cells::missing_arg(),
            VARIANT::from(formula),
        ],
    ))?;
    let _ = dispatch::call(&fc, "SetFirstPriority", &[]);
    let _ = dispatch::put(&fc, "StopIfTrue", VARIANT::from(false));
    Ok(fc)
}

fn tag_formula(mark: &str, test: &str) -> String {
    format!("=AND({test},LEN(\"{mark}\")>0)")
}

fn fill_color(fc: &IDispatch, color: i32) -> Result<(), XlError> {
    let interior = dispatch_of(dispatch::get(fc, "Interior"))?;
    dispatch::put(&interior, "Color", VARIANT::from(color))
        .map(drop)
        .map_err(com)
}

fn mark_cells(
    app: &IDispatch,
    r: &Region,
    addresses: &[String],
    mark: &str,
    color: i32,
) -> Result<(), XlError> {
    let sheet = dispatch_of(dispatch::get(&r.range, "Worksheet"))?;
    let mut union: Option<IDispatch> = None;
    for a in addresses.iter().take(2000) {
        let cell = dispatch_of(dispatch::get_with(
            &sheet,
            "Range",
            &[VARIANT::from(a.as_str())],
        ))?;
        union = Some(match union {
            None => cell,
            Some(u) => dispatch_of(dispatch::call(
                app,
                "Union",
                &[VARIANT::from(u), VARIANT::from(cell)],
            ))?,
        });
    }
    if let Some(u) = union {
        let fc = add_condition(&u, &tag_formula(mark, "TRUE"))?;
        fill_color(&fc, color)?;
    }
    Ok(())
}

/// Removes every conditional format of ours carrying `mark` (or all our marks if empty) from the
/// sheets of the active workbook.
fn clear_marks(app: &IDispatch, mark: &str) -> Result<usize, XlError> {
    let wb = dispatch_of(dispatch::get(app, "ActiveWorkbook"))?;
    let sheets = dispatch_of(dispatch::get(&wb, "Worksheets"))?;
    let n = count_of(&sheets, "Count")?;
    let mut removed = 0;
    for i in 1..=n {
        let Ok(ws) = dispatch_of(dispatch::get_with(&sheets, "Item", &[VARIANT::from(i)])) else {
            continue;
        };
        let Ok(cells) = dispatch_of(dispatch::get(&ws, "Cells")) else {
            continue;
        };
        let Ok(fcs) = dispatch_of(dispatch::get(&cells, "FormatConditions")) else {
            continue;
        };
        let Ok(count) = count_of(&fcs, "Count") else {
            continue;
        };
        for k in (1..=count).rev() {
            let Ok(fc) = dispatch_of(dispatch::get_with(&fcs, "Item", &[VARIANT::from(k)])) else {
                continue;
            };
            let formula = string_of(dispatch::get(&fc, "Formula1")).unwrap_or_default();
            let ours = if mark.is_empty() {
                formula.contains("\"xs-")
            } else {
                formula.contains(&format!("\"{mark}\""))
            };
            if ours && dispatch::call(&fc, "Delete", &[]).is_ok() {
                removed += 1;
            }
        }
    }
    Ok(removed)
}

/// The four edges of a region as conditional borders (the region's hidden name locates it).
fn outline(r: &Region) -> Result<(), XlError> {
    let name = format!("{}{}", regions::NAME_PREFIX, r.label);
    // xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
    let edges = [
        (-4131i32, format!("COLUMN()=MIN(COLUMN({name}))")),
        (-4160, format!("ROW()=MIN(ROW({name}))")),
        (-4107, format!("ROW()=MAX(ROW({name}))")),
        (-4152, format!("COLUMN()=MAX(COLUMN({name}))")),
    ];
    for (edge, test) in edges {
        let fc = add_condition(&r.range, &tag_formula(MARK_OUTLINE, &test))?;
        let borders = dispatch_of(dispatch::get(&fc, "Borders"))?;
        let b = dispatch_of(dispatch::get_with(&borders, "Item", &[VARIANT::from(edge)]))?;
        dispatch::put(&b, "LineStyle", VARIANT::from(1i32)).map_err(com)?;
        dispatch::put(&b, "Color", VARIANT::from(COLOR_OUTLINE)).map_err(com)?;
    }
    Ok(())
}

/// "Show": `anchors` highlights the anchor cells, `outlines` draws a frame around each region,
/// `clear` removes both (and the difference highlights) again.
pub(super) fn ui_show(app: &IDispatch, tag: &str) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        if tag == "clear" {
            let n = clear_marks(app, "").map_err(|e| e.message)?;
            return Ok(format!("Cleared {n} highlight(s) and outline(s)"));
        }
        let regs = all(&wb).map_err(|e| e.message)?;
        if regs.is_empty() {
            return Err(NO_REGIONS.into());
        }
        let _ = dispatch::put(app, "ScreenUpdating", VARIANT::from(false));
        let result = (|| -> Result<String, XlError> {
            match tag {
                "anchors" => {
                    clear_marks(app, MARK_ANCHOR)?;
                    for r in &regs {
                        let anchor = cell_address(&r.range, 0, 0);
                        mark_cells(app, r, &[anchor], MARK_ANCHOR, COLOR_ANCHOR)?;
                    }
                    Ok(format!(
                        "Highlighted the anchors of {}",
                        plural(regs.len(), "formula region", "formula regions")
                    ))
                }
                "outlines" => {
                    clear_marks(app, MARK_OUTLINE)?;
                    for r in &regs {
                        outline(r)?;
                    }
                    Ok(format!(
                        "Outlined {}",
                        plural(regs.len(), "formula region", "formula regions")
                    ))
                }
                _ => Err(XlError::new("unknown action")),
            }
        })();
        let _ = dispatch::put(app, "ScreenUpdating", VARIANT::from(true));
        result.map_err(|e| e.message)
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not mark the formula regions.", &e),
    }
}

/// "Roll": `tag` is `start:1`, `end:1`, `both:1`, or the same with -1.
pub(super) fn ui_roll(app: &IDispatch, tag: &str) {
    let go = || -> Result<String, String> {
        let (kind, n) = tag.split_once(':').ok_or("unknown roll action")?;
        let n: i32 = n.parse().map_err(|_| "unknown roll action".to_owned())?;
        let (start, end) = match kind {
            "start" => (n, 0),
            "end" => (0, n),
            "both" => (n, n),
            _ => return Err("unknown roll action".into()),
        };
        let wb = active_workbook(app)?;
        let r = region_here(app, &wb)?;
        let (rows, cols) = (r.info.rows as i32, r.info.cols as i32);
        let plan = regions::roll(rows, cols, default_axis(rows, cols), start, end)?;
        let leaving: i64 = plan
            .leaving
            .iter()
            .map(|x| i64::from(x.rows) * i64::from(x.cols))
            .sum();
        let mut old = OldCells::Keep;
        if leaving > 0 {
            let text = format!(
                "{leaving} cell(s) leave region '{}'.\n\nYes: turn them into values (the usual roll-forward)\nNo: keep their formulas\nCancel: do nothing",
                r.label
            );
            match ask3(excel_hwnd(app), &text, "Roll formula region") {
                Choice::Yes => old = OldCells::Values,
                Choice::No => {}
                Choice::Cancel => return Ok("Nothing changed".into()),
            }
        }
        match with_confirm(app, |o| roll(&r, start, end, old, o)).map_err(|e| e.message)? {
            Some(moved) => Ok(format!(
                "Formula region '{}' is now {}",
                moved.label, moved.info.address
            )),
            None => Ok("Nothing changed".into()),
        }
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not roll the formula region.", &e),
    }
}

// ------------------------------------------------------------------ rules and the workbook roll

fn rule_for_tag(tag: &str) -> Result<Rule, String> {
    let (word, old) = match tag.split_once(':') {
        Some((w, o)) => (w, OldCells::from_word(o).ok_or("unknown rule")?),
        None => (tag, OldCells::Values),
    };
    Ok(Rule {
        mode: RollMode::from_word(word).ok_or("unknown rule")?,
        old,
    })
}

/// A rule from the Define dropdown (`none`, `start`, `end`, `slide`; `start:keep`, `slide:clear`
/// ... say what happens to the cells that leave). Applied to the formula regions the selection
/// touches; a selection that touches none is defined as a new region with that rule.
pub(super) fn ui_rule(app: &IDispatch, tag: &str) {
    let go = || -> Result<String, String> {
        let rule = rule_for_tag(tag)?;
        let wb = active_workbook(app)?;
        let sel = dispatch_of(dispatch::get(app, "Selection"))
            .map_err(|_| "Select cells first.".to_owned())?;
        let mut n = 0;
        for r in all(&wb).map_err(|e| e.message)? {
            if overlaps(app, &sel, &r.range) {
                set_rule(&r, rule).map_err(|e| e.message)?;
                n += 1;
            }
        }
        if n > 0 {
            return Ok(format!("{n} formula region(s) now {}", rule.describe()));
        }
        match with_confirm(app, |o| define(&wb, &sel, None, o)).map_err(|e| e.message)? {
            Some((r, _)) => {
                set_rule(&r, rule).map_err(|e| e.message)?;
                Ok(format!(
                    "Formula region '{}' defined: {} ({})",
                    r.label,
                    r.info.address,
                    rule.describe()
                ))
            }
            None => Ok("Formula region not defined".to_owned()),
        }
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not set the roll rule.", &e),
    }
}

fn plural(n: usize, one: &str, many: &str) -> String {
    format!("{n} {}", if n == 1 { one } else { many })
}

/// "Roll all": every region that has a rule moves by `n` periods (`1` or `-1`), each by its own
/// rule. Regions blocked by hardcodes in their way are listed and rolled only if the user agrees.
pub(super) fn ui_roll_all(app: &IDispatch, tag: &str) {
    let go = || -> Result<String, String> {
        let n: i32 = tag.parse().map_err(|_| "unknown roll amount".to_owned())?;
        let wb = active_workbook(app)?;
        let regs: Vec<Region> = all(&wb)
            .map_err(|e| e.message)?
            .into_iter()
            .filter(|r| r.info.rule.mode != RollMode::None)
            .collect();
        if regs.is_empty() {
            return Err(
                "No formula region has a roll rule yet. Select regions and choose a rule under Rules."
                    .into(),
            );
        }
        let count = |m: RollMode| regs.iter().filter(|r| r.info.rule.mode == m).count();
        let text = format!(
            "Roll {} by {n:+}:\n\n  {} move their start\n  {} extend their end\n  {} slide their window\n\nThe regions without a rule stay as they are.",
            plural(regs.len(), "formula region", "formula regions"),
            count(RollMode::Start),
            count(RollMode::End),
            count(RollMode::Slide)
        );
        if !message(
            excel_hwnd(app),
            &text,
            "Roll all formula regions",
            MB_YESNO | MB_ICONINFORMATION,
        ) {
            return Ok("Nothing changed".into());
        }
        let overwrite = crate::settings::load().overwrite_regions;
        let _ = dispatch::put(app, "ScreenUpdating", VARIANT::from(false));
        let mut rolled = 0usize;
        let mut blocked: Vec<(Region, String)> = Vec::new();
        let mut failed: Vec<String> = Vec::new();
        for r in regs {
            let (start, end) = r.info.rule.steps(n);
            match roll(&r, start, end, r.info.rule.old, overwrite) {
                Ok(_) => rolled += 1,
                Err(e) if is_blocked(&e) => blocked.push((r, e.message)),
                Err(e) => failed.push(format!("{}: {}", r.label, e.message)),
            }
        }
        if !blocked.is_empty() {
            let names: Vec<&str> = blocked
                .iter()
                .take(15)
                .map(|(r, _)| r.label.as_str())
                .collect();
            let more = blocked.len().saturating_sub(15);
            let text = format!(
                "{} cannot roll because cells where they grow hold something other than their formula:\n\n  {}{}\n\nReplace those cells with the region's formula and roll them?",
                plural(blocked.len(), "region", "regions"),
                names.join("\n  "),
                if more > 0 {
                    format!("\n  ... and {more} more")
                } else {
                    String::new()
                }
            );
            if message(
                excel_hwnd(app),
                &text,
                "Roll all formula regions",
                MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2,
            ) {
                for (r, _) in blocked.drain(..) {
                    let (start, end) = r.info.rule.steps(n);
                    match roll(&r, start, end, r.info.rule.old, true) {
                        Ok(_) => rolled += 1,
                        Err(e) => failed.push(format!("{}: {}", r.label, e.message)),
                    }
                }
            }
        }
        let _ = dispatch::put(app, "ScreenUpdating", VARIANT::from(true));
        let mut summary = format!(
            "Rolled {}",
            plural(rolled, "formula region", "formula regions")
        );
        if !blocked.is_empty() {
            summary += &format!("; {} skipped", blocked.len());
        }
        if !failed.is_empty() {
            let shown: Vec<&str> = failed.iter().take(10).map(String::as_str).collect();
            fail(
                app,
                &format!(
                    "{} could not be rolled:",
                    plural(failed.len(), "region", "regions")
                ),
                &shown.join("\n"),
            );
        }
        Ok(summary)
    };
    let result = go();
    let _ = dispatch::put(app, "ScreenUpdating", VARIANT::from(true));
    match result {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not roll the formula regions.", &e),
    }
}

/// "Report": a sheet listing every region with its address, size and rule, to review them.
pub(super) fn ui_report(app: &IDispatch) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        let regs = all(&wb).map_err(|e| e.message)?;
        if regs.is_empty() {
            return Err("This workbook has no formula regions.".into());
        }
        let sheets = dispatch_of(dispatch::get(&wb, "Worksheets")).map_err(|e| e.message)?;
        let ws = match dispatch_of(dispatch::get_with(
            &sheets,
            "Item",
            &[VARIANT::from("Region rules")],
        )) {
            Ok(ws) => {
                let cells = dispatch_of(dispatch::get(&ws, "Cells")).map_err(|e| e.message)?;
                dispatch::call(&cells, "ClearContents", &[]).map_err(|e| com(e).message)?;
                ws
            }
            Err(_) => {
                let ws = dispatch_of(dispatch::call(&sheets, "Add", &[])).map_err(|e| e.message)?;
                dispatch::put(&ws, "Name", VARIANT::from("Region rules"))
                    .map_err(|e| com(e).message)?;
                ws
            }
        };
        let s = |t: &str| Value::Str(t.to_owned());
        let mut data = vec![
            s("Region"),
            s("Address"),
            s("Rows"),
            s("Columns"),
            s("Rule"),
            s("Old cells"),
        ];
        for r in &regs {
            data.push(s(&r.label));
            data.push(s(&r.info.address));
            data.push(Value::Num(f64::from(r.info.rows)));
            data.push(Value::Num(f64::from(r.info.cols)));
            data.push(s(r.info.rule.mode.word()));
            data.push(s(match r.info.rule.mode {
                RollMode::None => "",
                _ => r.info.rule.old.word(),
            }));
        }
        let table = Value::Array(crate::value::Array::new(regs.len() + 1, 6, data));
        let a1 = dispatch_of(dispatch::get_with(&ws, "Range", &[VARIANT::from("A1")]))
            .map_err(|e| e.message)?;
        let target = dispatch_of(dispatch::get_with(
            &a1,
            "Resize",
            &[VARIANT::from(regs.len() as i32 + 1), VARIANT::from(6i32)],
        ))
        .map_err(|e| e.message)?;
        dispatch::put(
            &target,
            "Value2",
            write_variant(&table).map_err(|e| e.message)?,
        )
        .map_err(|e| com(e).message)?;
        let cols = dispatch_of(dispatch::get(&ws, "Columns")).map_err(|e| e.message)?;
        let _ = dispatch::call(&cols, "AutoFit", &[]);
        let _ = dispatch::call(&ws, "Activate", &[]);
        Ok(format!(
            "Listed {} on the 'Region rules' sheet",
            plural(regs.len(), "formula region", "formula regions")
        ))
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not list the formula regions.", &e),
    }
}
