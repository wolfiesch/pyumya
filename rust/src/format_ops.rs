use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::str::FromStr;

use umya_spreadsheet::structs::{
    EnumTrait, HorizontalAlignmentValues, PatternValues, VerticalAlignmentValues,
};
use umya_spreadsheet::Spreadsheet;

use crate::utils::{a1_to_row_col, argb_to_hex, hex_to_argb, umya_border_style_to_str};

pub(crate) fn read_cell_format(
    book: &Spreadsheet,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<Py<PyAny>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let coord = (col0 + 1, row0 + 1);

    let d = PyDict::new(py);

    let cell = match ws.get_cell(coord) {
        Some(c) => c,
        None => return Ok(d.into_any().unbind()),
    };

    let style = cell.get_style();

    // Font properties
    if let Some(font) = style.get_font() {
        if *font.get_bold() {
            d.set_item("bold", true)?;
        }
        if *font.get_italic() {
            d.set_item("italic", true)?;
        }
        {
            let ul = font.get_underline();
            if !ul.is_empty() && ul != "none" {
                d.set_item("underline", ul.to_string())?;
            }
        }
        if *font.get_strikethrough() {
            d.set_item("strikethrough", true)?;
        }
        {
            let name = font.get_name();
            if !name.is_empty() {
                d.set_item("font_name", name.to_string())?;
            }
        }
        {
            let size = *font.get_size();
            if size > 0.0 {
                d.set_item("font_size", size)?;
            }
        }
        {
            let argb = font.get_color().get_argb();
            if !argb.is_empty() {
                let rgb = argb_to_hex(argb);
                if rgb != "000000" {
                    d.set_item("font_color", rgb)?;
                }
            }
        }
    }

    // Fill / background color
    if let Some(fill) = style.get_fill() {
        if let Some(pf) = fill.get_pattern_fill() {
            let p = pf.get_pattern_type().get_value_string();
            if !p.is_empty() && p != "none" {
                d.set_item("fill_type", p.to_string())?;
            }
            if let Some(fg) = pf.get_foreground_color() {
                let argb = fg.get_argb();
                if !argb.is_empty() {
                    let rgb = argb_to_hex(argb);
                    d.set_item("bg_color", rgb)?;
                }
            }
        }
    }

    // Number format
    if let Some(nf) = style.get_number_format() {
        let code = nf.get_format_code();
        if !code.is_empty() && code != "General" {
            d.set_item("number_format", code.to_string())?;
        }
    }

    // Alignment
    if let Some(align) = style.get_alignment() {
        let h = align.get_horizontal().get_value_string();
        if !h.is_empty() && h != "general" {
            d.set_item("h_align", h.to_string())?;
        }
        let v = align.get_vertical().get_value_string();
        if !v.is_empty() && v != "bottom" {
            d.set_item("v_align", v.to_string())?;
        }
        if *align.get_wrap_text() {
            d.set_item("wrap", true)?;
        }
        let rot = *align.get_text_rotation();
        if rot != 0 {
            d.set_item("rotation", rot)?;
        }
    }

    Ok(d.into_any().unbind())
}

pub(crate) fn write_cell_format(
    book: &mut Spreadsheet,
    sheet: &str,
    a1: &str,
    format_dict: &Bound<'_, PyAny>,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let dict = format_dict
        .cast::<PyDict>()
        .map_err(|_| PyErr::new::<PyValueError, _>("format_dict must be a dict"))?;

    let style = ws.get_style_mut(a1);

    // Font properties
    {
        let font = style.get_font_mut();

        if let Some(bold) = dict.get_item("bold")? {
            font.set_bold(bold.extract::<bool>()?);
        }
        if let Some(italic) = dict.get_item("italic")? {
            font.set_italic(italic.extract::<bool>()?);
        }
        if let Some(ul) = dict.get_item("underline")? {
            font.set_underline(ul.extract::<String>()?);
        }
        if let Some(st) = dict.get_item("strikethrough")? {
            font.set_strikethrough(st.extract::<bool>()?);
        }
        if let Some(name) = dict.get_item("font_name")? {
            font.set_name(name.extract::<String>()?);
        }
        if let Some(size) = dict.get_item("font_size")? {
            font.set_size(size.extract::<f64>()?);
        }
        if let Some(color) = dict.get_item("font_color")? {
            let c = color.extract::<String>()?;
            font.get_color_mut().set_argb(hex_to_argb(&c));
        }
    }

    // Fill / background via pattern fill
    if let Some(fill_type) = dict.get_item("fill_type")? {
        let ft = fill_type.extract::<String>()?;
        if ft.eq_ignore_ascii_case("none") {
            style
                .get_fill_mut()
                .get_pattern_fill_mut()
                .set_pattern_type(PatternValues::None);
        }
    }
    if let Some(bg) = dict.get_item("bg_color")? {
        let bg = bg.extract::<String>()?;
        let fill = style.get_fill_mut();
        let pf = fill.get_pattern_fill_mut();
        pf.set_pattern_type(PatternValues::Solid);
        pf.get_foreground_color_mut().set_argb(hex_to_argb(&bg));
    }

    // Number format
    if let Some(nf) = dict.get_item("number_format")? {
        style
            .get_number_format_mut()
            .set_format_code(nf.extract::<String>()?);
    }

    // Alignment
    if let Some(h) = dict.get_item("h_align")? {
        let h = h.extract::<String>()?;
        if let Ok(ha) = HorizontalAlignmentValues::from_str(&h) {
            style.get_alignment_mut().set_horizontal(ha);
        }
    }
    if let Some(v) = dict.get_item("v_align")? {
        let v = v.extract::<String>()?;
        if let Ok(va) = VerticalAlignmentValues::from_str(&v) {
            style.get_alignment_mut().set_vertical(va);
        }
    }
    if let Some(wrap) = dict.get_item("wrap")? {
        style
            .get_alignment_mut()
            .set_wrap_text(wrap.extract::<bool>()?);
    }
    if let Some(rot) = dict.get_item("rotation")? {
        let rot = rot.extract::<i64>()?;
        if let Ok(r) = u32::try_from(rot) {
            style.get_alignment_mut().set_text_rotation(r);
        }
    }

    Ok(())
}

pub(crate) fn read_cell_border(
    book: &Spreadsheet,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<Py<PyAny>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let coord = (col0 + 1, row0 + 1);

    let d = PyDict::new(py);

    let cell = match ws.get_cell(coord) {
        Some(c) => c,
        None => return Ok(d.into_any().unbind()),
    };

    let style = cell.get_style();
    if let Some(borders) = style.get_borders() {
        // Helper: read one edge. Returns None if style is "none" or empty.
        let read_edge = |e: &umya_spreadsheet::structs::Border| -> Option<(String, String)> {
            let style_str = e.get_border_style();
            if style_str.is_empty() || style_str == "none" {
                return None;
            }
            let argb = e.get_color().get_argb();
            let color_str = if argb.is_empty() {
                "000000".to_string()
            } else {
                argb_to_hex(argb)
            };
            Some((umya_border_style_to_str(style_str).to_string(), color_str))
        };

        if let Some((s, c)) = read_edge(borders.get_top()) {
            let edge = PyDict::new(py);
            edge.set_item("style", s)?;
            edge.set_item("color", c)?;
            d.set_item("top", edge)?;
        }
        if let Some((s, c)) = read_edge(borders.get_bottom()) {
            let edge = PyDict::new(py);
            edge.set_item("style", s)?;
            edge.set_item("color", c)?;
            d.set_item("bottom", edge)?;
        }
        if let Some((s, c)) = read_edge(borders.get_left()) {
            let edge = PyDict::new(py);
            edge.set_item("style", s)?;
            edge.set_item("color", c)?;
            d.set_item("left", edge)?;
        }
        if let Some((s, c)) = read_edge(borders.get_right()) {
            let edge = PyDict::new(py);
            edge.set_item("style", s)?;
            edge.set_item("color", c)?;
            d.set_item("right", edge)?;
        }
        if let Some((s, c)) = read_edge(borders.get_diagonal()) {
            let edge = PyDict::new(py);
            edge.set_item("style", s)?;
            edge.set_item("color", c)?;
            // umya doesn't distinguish up/down diagonal - use diagonal_up.
            d.set_item("diagonal_up", edge)?;
        }
    }

    Ok(d.into_any().unbind())
}

pub(crate) fn write_cell_border(
    book: &mut Spreadsheet,
    sheet: &str,
    a1: &str,
    border_dict: &Bound<'_, PyAny>,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let dict = border_dict
        .cast::<PyDict>()
        .map_err(|_| PyErr::new::<PyValueError, _>("border_dict must be a dict"))?;

    let style = ws.get_style_mut(a1);
    let borders = style.get_borders_mut();

    fn apply_edge(
        edge: &mut umya_spreadsheet::structs::Border,
        sub: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        if let Some(s) = sub.get_item("style")? {
            edge.set_border_style(s.extract::<String>()?);
        }
        if let Some(c) = sub.get_item("color")? {
            let c = c.extract::<String>()?;
            edge.get_color_mut().set_argb(hex_to_argb(&c));
        }
        Ok(())
    }

    if let Some(sub) = dict.get_item("top")? {
        if let Ok(d) = sub.cast::<PyDict>() {
            apply_edge(borders.get_top_mut(), &d)?;
        }
    }
    if let Some(sub) = dict.get_item("bottom")? {
        if let Ok(d) = sub.cast::<PyDict>() {
            apply_edge(borders.get_bottom_mut(), &d)?;
        }
    }
    if let Some(sub) = dict.get_item("left")? {
        if let Ok(d) = sub.cast::<PyDict>() {
            apply_edge(borders.get_left_mut(), &d)?;
        }
    }
    if let Some(sub) = dict.get_item("right")? {
        if let Ok(d) = sub.cast::<PyDict>() {
            apply_edge(borders.get_right_mut(), &d)?;
        }
    }
    if let Some(sub) = dict.get_item("diagonal_up")? {
        if let Ok(d) = sub.cast::<PyDict>() {
            apply_edge(borders.get_diagonal_mut(), &d)?;
        }
    }
    if let Some(sub) = dict.get_item("diagonal_down")? {
        if let Ok(d) = sub.cast::<PyDict>() {
            apply_edge(borders.get_diagonal_mut(), &d)?;
        }
    }

    Ok(())
}
