use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;

use umya_spreadsheet::Spreadsheet;

pub(crate) fn sheet_max_row(book: &Spreadsheet, sheet: &str) -> PyResult<u32> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
    Ok(ws.get_highest_row())
}

pub(crate) fn sheet_max_column(book: &Spreadsheet, sheet: &str) -> PyResult<u32> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
    Ok(ws.get_highest_column())
}
