use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::Hyperlink;
use umya_spreadsheet::Spreadsheet;

pub(crate) fn read_hyperlinks(
    book: &Spreadsheet,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<Py<PyAny>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let list = PyList::empty(py);

    for cell in ws.get_cell_collection_sorted() {
        let Some(link) = cell.get_hyperlink() else {
            continue;
        };

        let d = PyDict::new(py);
        d.set_item("cell", cell.get_coordinate().to_string())?;

        let target = link.get_url();
        if !target.is_empty() {
            // Some relationship targets are stored as XML attribute values and
            // can retain entity-escaping (e.g. "&amp;") depending on the reader.
            // Normalize the common case so ExcelBench comparisons match.
            let normalized = target.replace("&amp;", "&");
            d.set_item("target", normalized)?;
        }

        let tooltip = link.get_tooltip();
        if !tooltip.is_empty() {
            d.set_item("tooltip", tooltip.to_string())?;
        }

        d.set_item("internal", *link.get_location())?;
        list.append(d)?;
    }

    Ok(list.into_any().unbind())
}

pub(crate) fn add_hyperlink(
    book: &mut Spreadsheet,
    sheet: &str,
    a1: &str,
    target: &str,
    tooltip: Option<&str>,
    internal: bool,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let mut link = Hyperlink::default();
    link.set_url(target);
    if let Some(tt) = tooltip {
        if !tt.is_empty() {
            link.set_tooltip(tt);
        }
    }
    link.set_location(internal);

    ws.get_cell_mut(a1).set_hyperlink(link);
    Ok(())
}
