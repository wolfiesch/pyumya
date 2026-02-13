use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::drawing::spreadsheet::MarkerType;
use umya_spreadsheet::structs::Image;
use umya_spreadsheet::Spreadsheet;

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
        marker.set_col_off(x * EMU_PER_PX);
        marker.set_row_off(y * EMU_PER_PX);
    }

    let mut img = Image::default();
    img.new_image(path, marker);
    ws.add_image(img);
    Ok(())
}
