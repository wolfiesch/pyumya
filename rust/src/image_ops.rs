use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::drawing::spreadsheet::MarkerType;
use umya_spreadsheet::structs::Image;
use umya_spreadsheet::Spreadsheet;

// EMUs (English Metric Units) per CSS pixel, assuming the 96 DPI model used by
// Office Open XML: 914_400 EMU/inch / 96 px/inch = 9_525 EMU/px.
// This does not account for high-DPI scaling; it follows OOXML's 96 DPI assumption.
const EMU_PER_PX: i32 = 9525;

pub(crate) fn read_images(book: &Spreadsheet, py: Python<'_>, sheet: &str) -> PyResult<Py<PyAny>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    // umya-spreadsheet defers parsing drawings (including images) until the
    // worksheet is deserialized. Touch the cell collection to force
    // deserialization so images are populated when reading existing files.
    let _ = ws.get_cell_collection_sorted();

    let list = PyList::empty(py);
    for img in ws.get_image_collection() {
        let d = PyDict::new(py);

        d.set_item("cell", img.get_coordinate())?;

        let name = img.get_image_name();
        if !name.is_empty() {
            d.set_item("path", format!("/xl/media/{name}"))?;
        }

        let anchor = if img.get_two_cell_anchor().is_some() {
            "twoCell"
        } else {
            "oneCell"
        };
        d.set_item("anchor", anchor)?;

        let marker = img.get_from_marker_type();
        let x = (*marker.get_col_off() / EMU_PER_PX) as i32;
        let y = (*marker.get_row_off() / EMU_PER_PX) as i32;
        if x != 0 || y != 0 {
            d.set_item("offset", vec![x, y])?;
        }

        list.append(d)?;
    }

    Ok(list.into_any().unbind())
}

pub(crate) fn add_image(
    book: &mut Spreadsheet,
    sheet: &str,
    cell: &str,
    path: &str,
    offset: Option<(i32, i32)>,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let mut marker = MarkerType::default();
    marker.set_coordinate(cell);
    if let Some((x, y)) = offset {
        let col_off = x
            .checked_mul(EMU_PER_PX)
            .ok_or_else(|| PyErr::new::<PyValueError, _>("Image column offset too large (overflow)"))?;
        let row_off = y
            .checked_mul(EMU_PER_PX)
            .ok_or_else(|| PyErr::new::<PyValueError, _>("Image row offset too large (overflow)"))?;
        marker.set_col_off(col_off);
        marker.set_row_off(row_off);
    }

    let mut img = Image::default();
    if !std::path::Path::new(path).exists() {
        return Err(PyErr::new::<PyValueError, _>(format!("Image file not found: {path}")));
    }
    img.new_image(path, marker);
    ws.add_image(img);
    Ok(())
}
