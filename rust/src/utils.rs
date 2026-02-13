use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use pyo3::prelude::*;
use pyo3::types::PyDict;

// ---------------------------------------------------------------------------
// A1 helpers
// ---------------------------------------------------------------------------

/// Convert an A1-style cell reference (e.g. "B3") to (row0, col0) 0-based.
pub fn a1_to_row_col(a1: &str) -> Result<(u32, u32), String> {
    let mut col: u32 = 0;
    let mut row_digits = String::new();

    for ch in a1.chars() {
        if ch.is_ascii_alphabetic() {
            let uc = ch.to_ascii_uppercase() as u8;
            let val = (uc - b'A' + 1) as u32;
            col = col * 26 + val;
        } else if ch.is_ascii_digit() {
            row_digits.push(ch);
        } else {
            return Err(format!("Invalid cell reference: {a1}"));
        }
    }

    if col == 0 || row_digits.is_empty() {
        return Err(format!("Invalid cell reference: {a1}"));
    }

    let row_1: u32 = row_digits
        .parse()
        .map_err(|_| format!("Invalid cell reference: {a1}"))?;
    if row_1 == 0 {
        return Err(format!("Invalid cell reference: {a1}"));
    }

    Ok((row_1 - 1, col - 1))
}

/// Convert a column letter (e.g. "A", "AA") into a 1-based column index.
pub fn col_letter_to_u32(col_str: &str) -> Result<u32, String> {
    let mut col: u32 = 0;
    for ch in col_str.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(format!("Invalid column string: {col_str}"));
        }
        let uc = ch.to_ascii_uppercase() as u8;
        col = col * 26 + (uc - b'A' + 1) as u32;
    }
    if col == 0 {
        return Err(format!("Invalid column string: {col_str}"));
    }
    Ok(col)
}

/// Convert a 1-based column index into a column letter.
pub fn u32_to_col_letter(col: u32) -> String {
    if col == 0 {
        return String::new();
    }
    let mut n = col;
    let mut out: Vec<u8> = Vec::new();
    while n > 0 {
        let rem = ((n - 1) % 26) as u8;
        out.push(b'A' + rem);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap_or_default()
}

// ---------------------------------------------------------------------------
// Color helpers: ARGB <-> hex
// ---------------------------------------------------------------------------

/// Convert ARGB "FFRRGGBB" or "RRGGBB" to "#RRGGBB".
pub fn argb_to_hex(argb: &str) -> String {
    let s = argb.trim();
    if s.len() == 8 {
        format!("#{}", &s[2..])
    } else if s.len() == 6 {
        format!("#{s}")
    } else if s.starts_with('#') {
        s.to_string()
    } else {
        format!("#{s}")
    }
}

/// Convert "#RRGGBB" (or "RRGGBB") to "FFRRGGBB" ARGB.
pub fn hex_to_argb(hex: &str) -> String {
    let s = hex.strip_prefix('#').unwrap_or(hex).trim();
    if s.len() == 8 {
        // Already ARGB.
        s.to_string()
    } else {
        format!("FF{s}")
    }
}

// ---------------------------------------------------------------------------
// Date helpers
// ---------------------------------------------------------------------------

pub fn looks_like_date_format(code: &str) -> bool {
    // Heuristic: date formats typically include year + day tokens.
    let lc = code.to_ascii_lowercase();
    lc.contains('y') && lc.contains('d')
}

pub fn excel_serial_to_naive_datetime(serial: f64) -> Option<NaiveDateTime> {
    // Excel 1900 date system, with the standard 1900 leap-year bug adjustment.
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_time(NaiveTime::MIN);
    let mut f = serial;
    if f < 60.0 {
        f += 1.0;
    }
    let total_ms = (f * 86_400_000.0).round() as i64;
    epoch.checked_add_signed(Duration::milliseconds(total_ms))
}

pub fn naive_datetime_to_excel_serial(dt: NaiveDateTime) -> Option<f64> {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_time(NaiveTime::MIN);
    let delta = dt - epoch;
    let total_ms = delta.num_milliseconds();
    Some(total_ms as f64 / 86_400_000.0)
}

pub fn parse_iso_date(s: &str) -> Option<NaiveDate> {
    NaiveDate::parse_from_str(s, "%Y-%m-%d").ok()
}

pub fn parse_iso_datetime(s: &str) -> Option<NaiveDateTime> {
    let raw = s.trim_end_matches('Z');
    NaiveDateTime::parse_from_str(raw, "%Y-%m-%dT%H:%M:%S")
        .ok()
        .or_else(|| NaiveDateTime::parse_from_str(raw, "%Y-%m-%dT%H:%M:%S%.f").ok())
}

// ---------------------------------------------------------------------------
// Py helpers
// ---------------------------------------------------------------------------

pub(crate) fn cell_blank(py: Python<'_>) -> PyResult<Py<PyAny>> {
    let d = PyDict::new(py);
    // The Python layer treats missing "value" as blank.
    d.set_item("type", "blank")?;
    Ok(d.into_any().unbind())
}

pub(crate) fn cell_with_value<'py, V>(py: Python<'py>, t: &str, value: V) -> PyResult<Py<PyAny>>
where
    V: IntoPyObject<'py>,
{
    let d = PyDict::new(py);
    d.set_item("type", t)?;
    d.set_item("value", value)?;
    Ok(d.into_any().unbind())
}

// ---------------------------------------------------------------------------
// Border helpers
// ---------------------------------------------------------------------------

/// Map umya border style string to canonical style names.
pub fn umya_border_style_to_str(style: &str) -> &'static str {
    match style.to_ascii_lowercase().as_str() {
        "thin" => "thin",
        "medium" => "medium",
        "thick" => "thick",
        "double" => "double",
        "dashed" => "dashed",
        "dotted" => "dotted",
        "hair" => "hair",
        "mediumdashed" => "mediumDashed",
        "dashdot" => "dashDot",
        "mediumdashdot" => "mediumDashDot",
        "dashdotdot" => "dashDotDot",
        "mediumdashdotdot" => "mediumDashDotDot",
        "slantdashdot" => "slantDashDot",
        _ => "none",
    }
}
