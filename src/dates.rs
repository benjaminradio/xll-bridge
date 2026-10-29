//! Excel date numbers <-> ISO 8601 text.
//!
//! Excel stores a date as a day count (the 1900 system: 1 = 1900-01-01) plus a fraction of a day
//! for the time. Dates stay numbers everywhere in this add-in (they sort and subtract correctly);
//! these helpers only convert at the edges. Serial numbers from 61 on map exactly onto the
//! calendar (day 0 is 1899-12-30). Excel's phantom 1900-02-29 (serial 60) and the 1904 date system
//! are not modelled.

/// Days from 1970-01-01 to 1899-12-30, Excel's day 0.
const EXCEL_EPOCH_FROM_UNIX_DAYS: i64 = -25_569;
const MS_PER_DAY: i64 = 86_400_000;

/// Days since 1970-01-01 -> (year, month, day), proleptic Gregorian (Howard Hinnant's algorithm).
fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
}

fn days_from_civil(y: i64, m: u32, d: u32) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = y.div_euclid(400);
    let yoe = y.rem_euclid(400);
    let mp = if m > 2 { m - 3 } else { m + 9 } as i64;
    let doy = (153 * mp + 2) / 5 + i64::from(d) - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

fn days_in_month(y: i64, m: u32) -> u32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        _ if (y % 4 == 0 && y % 100 != 0) || y % 400 == 0 => 29,
        _ => 28,
    }
}

/// `2026-10-29` for whole days, `2026-10-29T13:45:00` (plus `.mmm` when needed) otherwise.
/// `None` for numbers outside the years 1900..9999 or not finite.
pub fn serial_to_iso(serial: f64) -> Option<String> {
    if !serial.is_finite() || !(61.0..2_958_466.0).contains(&serial) {
        return None;
    }
    let total_ms = (serial * MS_PER_DAY as f64).round() as i64;
    let days = total_ms.div_euclid(MS_PER_DAY);
    let ms = total_ms.rem_euclid(MS_PER_DAY);
    let (y, m, d) = civil_from_days(days + EXCEL_EPOCH_FROM_UNIX_DAYS);
    let mut out = format!("{y:04}-{m:02}-{d:02}");
    if ms != 0 {
        let (h, mi, s, milli) = (ms / 3_600_000, ms / 60_000 % 60, ms / 1000 % 60, ms % 1000);
        out.push_str(&format!("T{h:02}:{mi:02}:{s:02}"));
        if milli != 0 {
            out.push_str(&format!(".{milli:03}"));
        }
    }
    Some(out)
}

fn digits(s: &str) -> Option<i64> {
    (!s.is_empty() && s.bytes().all(|b| b.is_ascii_digit()))
        .then(|| s.parse().ok())
        .flatten()
}

/// Parses `YYYY-MM-DD`, optionally followed by `T` or a space and `HH:MM[:SS[.fff]]`, optionally
/// ending in `Z` (ignored: no time zone arithmetic is done). Returns the Excel serial number.
pub fn iso_to_serial(text: &str) -> Option<f64> {
    let t = text.trim().trim_end_matches(['Z', 'z']);
    let (date, time) = match t.split_once(['T', 't', ' ']) {
        Some((d, tm)) => (d, Some(tm.trim())),
        None => (t, None),
    };
    let mut parts = date.split('-');
    let (y, m, d) = (
        digits(parts.next()?)?,
        digits(parts.next()?)?,
        digits(parts.next()?)?,
    );
    if parts.next().is_some() || !(1900..=9999).contains(&y) || !(1..=12).contains(&m) {
        return None;
    }
    let (m, d) = (m as u32, d as u32);
    if d < 1 || d > days_in_month(y, m) {
        return None;
    }
    let mut ms = 0i64;
    if let Some(tm) = time {
        let (hms, frac) = match tm.split_once('.') {
            Some((a, f)) => (a, Some(f)),
            None => (tm, None),
        };
        let mut p = hms.split(':');
        let h = digits(p.next()?)?;
        let mi = digits(p.next()?)?;
        let s = p.next().map_or(Some(0), digits)?;
        if p.next().is_some() || h > 23 || mi > 59 || s > 59 {
            return None;
        }
        ms = (h * 3600 + mi * 60 + s) * 1000;
        if let Some(f) = frac {
            let f3: String = f.chars().chain("000".chars()).take(3).collect();
            ms += digits(&f3)?;
        }
    }
    let days = days_from_civil(y, m, d) - EXCEL_EPOCH_FROM_UNIX_DAYS;
    if days < 61 {
        return None; // before 1900-03-01: Excel's calendar differs there
    }
    Some((days * MS_PER_DAY + ms) as f64 / MS_PER_DAY as f64)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn known_dates() {
        assert_eq!(serial_to_iso(61.0).as_deref(), Some("1900-03-01"));
        assert_eq!(serial_to_iso(25_569.0).as_deref(), Some("1970-01-01"));
        assert_eq!(serial_to_iso(36_526.0).as_deref(), Some("2000-01-01"));
        assert_eq!(serial_to_iso(45_292.0).as_deref(), Some("2024-01-01"));
        assert_eq!(serial_to_iso(45_351.0).as_deref(), Some("2024-02-29"));
        assert_eq!(serial_to_iso(2_958_465.0).as_deref(), Some("9999-12-31"));
    }

    #[test]
    fn times_and_fractions() {
        assert_eq!(
            serial_to_iso(45_292.5).as_deref(),
            Some("2024-01-01T12:00:00")
        );
        assert_eq!(
            serial_to_iso(45_292.0 + 1.0 / 86_400.0).as_deref(),
            Some("2024-01-01T00:00:01")
        );
        assert_eq!(
            serial_to_iso(45_292.0 + 1.0 / 86_400_000.0 * 250.0).as_deref(),
            Some("2024-01-01T00:00:00.250")
        );
    }

    #[test]
    fn out_of_range_and_garbage() {
        assert_eq!(serial_to_iso(0.0), None);
        assert_eq!(serial_to_iso(60.0), None);
        assert_eq!(serial_to_iso(f64::NAN), None);
        assert_eq!(serial_to_iso(3_000_000.0), None);
        for bad in [
            "",
            "2024",
            "2024-13-01",
            "2023-02-29",
            "2024-01-01T25:00",
            "2024-1-1x",
            "1899-12-31",
        ] {
            assert_eq!(iso_to_serial(bad), None, "{bad}");
        }
    }

    #[test]
    fn parsing_round_trips() {
        assert_eq!(iso_to_serial("2024-02-29"), Some(45_351.0));
        assert_eq!(iso_to_serial("2024-01-01T12:00"), Some(45_292.5));
        assert_eq!(iso_to_serial("2024-01-01 12:00:00Z"), Some(45_292.5));
        for serial in [
            61.0,
            36_526.25,
            45_351.75,
            45_292.0 + 3.0 / 86_400.0,
            2_958_465.0,
        ] {
            let iso = serial_to_iso(serial).unwrap();
            let back = iso_to_serial(&iso).unwrap();
            assert!((back - serial).abs() < 1e-9, "{iso}: {back} vs {serial}");
        }
    }
}
