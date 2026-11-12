//! Formula regions in Excel: the object-model side of [`crate::regions`] and the ribbon actions.
//!
//! A region is a hidden workbook name (`_xs_<label>`) that Excel keeps pointing at the right
//! cells through inserts and deletes; the formula itself is the top-left (anchor) cell's own.

use std::path::Path;

use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;

use super::cells::{
    com, count_of, describe, dispatch_of, open_book, read_variant, resolve, string_of,
};
use super::dispatch;
use crate::regions::{
    self, Axis, CheckReport, DevKind, Rect, RegionInfo, default_axis, label_of, name_for,
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
        out.push(Region {
            label: label.to_owned(),
            name: nm,
            range,
            info: RegionInfo {
                name: label.to_owned(),
                address: info.address,
                rows: info.rows,
                cols: info.cols,
            },
        });
    }
    Ok(out)
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
                "the range overlaps the formula region '{}'",
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
        },
        label,
        name: nm,
        range: range.clone(),
    })
}

/// Makes `range` a region and fills it from its anchor; on failure nothing is left behind.
pub(super) fn define(
    wb: &IDispatch,
    range: &IDispatch,
    label: Option<String>,
    overwrite: bool,
) -> Result<Region, XlError> {
    let region = create(wb, range, label)?;
    if let Err(e) = apply(&region, overwrite) {
        let _ = dispatch::call(&region.name, "Delete", &[]);
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
    freeze: bool,
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
    if freeze {
        for rect in &plan.leaving {
            let rng = rect_range(&top_left, *rect)?;
            let v = dispatch::get(&rng, "Value2").map_err(com)?;
            dispatch::put(&rng, "Value2", v).map_err(com)?;
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
            Ok(one(&define(&wb, &range, label, overwrite)?))
        }
        RegionOp::Roll {
            name,
            start,
            end,
            freeze,
            overwrite,
        } => {
            let r = find(&wb, &name)?;
            Ok(one(&roll(&r, start, end, freeze, overwrite)?))
        }
        RegionOp::Remove { name } => {
            remove(&find(&wb, &name)?)?;
            Ok(Reply::Done)
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

/// "Define": the selection becomes a region.
pub(super) fn ui_define(app: &IDispatch) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        let sel = dispatch_of(dispatch::get(app, "Selection")).map_err(|_| {
            "Select the cells of the region first (the formula goes in its top-left cell)."
                .to_owned()
        })?;
        match with_confirm(app, |o| define(&wb, &sel, None, o)).map_err(|e| e.message)? {
            Some(r) => Ok(format!(
                "Formula region '{}' defined: {} ({} cells)",
                r.label,
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

/// "Remove": forget the region around the active cell; the cells keep their formulas.
pub(super) fn ui_remove(app: &IDispatch) {
    let go = || -> Result<String, String> {
        let wb = active_workbook(app)?;
        let r = region_here(app, &wb)?;
        remove(&r).map_err(|e| e.message)?;
        Ok(format!(
            "Formula region '{}' removed; its cells keep their formulas",
            r.label
        ))
    };
    match go() {
        Ok(t) => set_status_bar(app, &t),
        Err(e) => fail(app, "Could not remove the formula region.", &e),
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

/// "Check": every region of the active workbook against its anchor.
pub(super) fn ui_check(app: &IDispatch) {
    let go = || -> Result<(), String> {
        let wb = active_workbook(app)?;
        let regions = all(&wb).map_err(|e| e.message)?;
        if regions.is_empty() {
            return Err("There are no formula regions in this workbook. Type the formula in the top-left cell of a block, select the block and use Define.".into());
        }
        let mut reports = Vec::new();
        for r in &regions {
            match check(r) {
                Ok(rep) => reports.push((r, Ok(rep))),
                Err(e) => reports.push((r, Err(e.message))),
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
    };
    if let Err(e) = go() {
        fail(app, "Formula regions", &e);
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
        let mut freeze = false;
        if leaving > 0 {
            let text = format!(
                "{leaving} cell(s) leave region '{}'.\n\nYes: turn them into values (the usual roll-forward)\nNo: keep their formulas\nCancel: do nothing",
                r.label
            );
            match ask3(excel_hwnd(app), &text, "Roll formula region") {
                Choice::Yes => freeze = true,
                Choice::No => {}
                Choice::Cancel => return Ok("Nothing changed".into()),
            }
        }
        match with_confirm(app, |o| roll(&r, start, end, freeze, o)).map_err(|e| e.message)? {
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
