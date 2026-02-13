use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;

use umya_spreadsheet::structs::{Coordinate, Pane, PaneStateValues, PaneValues, SheetView};
use umya_spreadsheet::Spreadsheet;

use crate::utils::{a1_to_row_col, col_letter_to_u32};

pub(crate) fn read_row_height(book: &Spreadsheet, sheet: &str, row: u32) -> PyResult<Option<f64>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    if row == 0 {
        return Err(PyErr::new::<PyValueError, _>("row must be >= 1"));
    }

    if let Some(rd) = ws.get_row_dimension(&row) {
        let h = rd.get_height();
        if h > &0.0 {
            return Ok(Some(*h));
        }
    }
    Ok(None)
}

pub(crate) fn set_row_height(
    book: &mut Spreadsheet,
    sheet: &str,
    row: u32,
    height: f64,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    if row == 0 {
        return Err(PyErr::new::<PyValueError, _>("row must be >= 1"));
    }

    ws.get_row_dimension_mut(&row).set_height(height);
    Ok(())
}

pub(crate) fn read_column_width(
    book: &Spreadsheet,
    sheet: &str,
    col_str: &str,
) -> PyResult<Option<f64>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let col_idx = col_letter_to_u32(col_str).map_err(|e| PyErr::new::<PyValueError, _>(e))?;
    if let Some(cd) = ws.get_column_dimension_by_number(&col_idx) {
        let w = cd.get_width();
        if w > &0.0 {
            return Ok(Some(*w));
        }
    }
    Ok(None)
}

pub(crate) fn set_column_width(
    book: &mut Spreadsheet,
    sheet: &str,
    col_str: &str,
    width: f64,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let col_idx = col_letter_to_u32(col_str).map_err(|e| PyErr::new::<PyValueError, _>(e))?;
    ws.get_column_dimension_by_number_mut(&col_idx)
        .set_width(width);
    Ok(())
}

pub(crate) fn merge_cells(book: &mut Spreadsheet, sheet: &str, range_str: &str) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    ws.add_merge_cells(range_str);
    Ok(())
}

pub(crate) fn unmerge_cells(book: &mut Spreadsheet, sheet: &str, range_str: &str) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let target = range_str.trim().to_ascii_uppercase();
    ws.get_merge_cells_mut()
        .retain(|r| r.get_range().to_ascii_uppercase() != target);
    Ok(())
}

pub(crate) fn get_merged_ranges(book: &Spreadsheet, sheet: &str) -> PyResult<Vec<String>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    Ok(ws.get_merge_cells().iter().map(|r| r.get_range()).collect())
}

pub(crate) fn get_freeze_panes(book: &Spreadsheet, sheet: &str) -> PyResult<Option<String>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    for sv in ws.get_sheets_views().get_sheet_view_list() {
        if let Some(pane) = sv.get_pane() {
            let coord = pane.get_top_left_cell().to_string();
            if coord.is_empty() || coord == "A1" {
                return Ok(None);
            }
            return Ok(Some(coord));
        }
    }
    Ok(None)
}

pub(crate) fn set_freeze_panes(
    book: &mut Spreadsheet,
    sheet: &str,
    a1: Option<&str>,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let Some(a1) = a1.map(str::trim).filter(|s| !s.is_empty()) else {
        ws.get_sheet_views_mut().get_sheet_view_list_mut().clear();
        return Ok(());
    };

    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row1 = row0 + 1;
    let col1 = col0 + 1;

    let x_split = if col1 > 1 { (col1 - 1) as f64 } else { 0.0 };
    let y_split = if row1 > 1 { (row1 - 1) as f64 } else { 0.0 };

    // A1 (or equivalents) means "no freeze".
    if x_split == 0.0 && y_split == 0.0 {
        ws.get_sheet_views_mut().get_sheet_view_list_mut().clear();
        return Ok(());
    }

    let mut pane = Pane::default();
    if x_split > 0.0 {
        pane.set_horizontal_split(x_split);
    }
    if y_split > 0.0 {
        pane.set_vertical_split(y_split);
    }
    let mut coord = Coordinate::default();
    coord.set_coordinate(a1);
    pane.set_top_left_cell(coord);
    pane.set_state(PaneStateValues::Frozen);

    // Avoid PaneValues::TopRight due to upstream casing bug; BottomRight works
    // for vertical splits as well.
    if y_split > 0.0 && x_split == 0.0 {
        pane.set_active_pane(PaneValues::BottomLeft);
    } else {
        pane.set_active_pane(PaneValues::BottomRight);
    }

    let sv_list = ws.get_sheet_views_mut().get_sheet_view_list_mut();
    if sv_list.is_empty() {
        sv_list.push(SheetView::default());
    }
    sv_list[0].set_pane(pane);
    Ok(())
}
