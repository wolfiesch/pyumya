use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;

use std::path::Path;

use umya_spreadsheet::{new_file, reader, writer, Spreadsheet};

use crate::{cell_ops, worksheet};

/// Low-level Rust workbook handle wrapping umya-spreadsheet.
///
/// This is the internal FFI class. The Python `Workbook` class wraps this
/// and provides the user-facing API.
#[pyclass(unsendable)]
pub struct RustWorkbook {
    book: Spreadsheet,
}

#[pymethods]
impl RustWorkbook {
    #[new]
    pub fn new() -> Self {
        let mut book = new_file();
        let _ = book.remove_sheet_by_name("Sheet1");
        Self { book }
    }

    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let p = Path::new(path);
        let book = reader::xlsx::read(p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open: {e}")))?;
        Ok(Self { book })
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.book
            .get_sheet_collection()
            .iter()
            .map(|s| s.get_name().to_string())
            .collect()
    }

    pub fn sheet_count(&self) -> usize {
        self.book.get_sheet_collection().iter().count()
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        self.book
            .new_sheet(name)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("{e}")))?;
        Ok(())
    }

    pub fn remove_sheet(&mut self, name: &str) -> PyResult<()> {
        self.book
            .remove_sheet_by_name(name)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("{e}")))?;
        Ok(())
    }

    // =========================================================================
    // Phase 1: Core cell R/W
    // =========================================================================

    pub fn read_cell_value(&self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<Py<PyAny>> {
        cell_ops::read_cell_value(&self.book, py, sheet, a1)
    }

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        cell_ops::write_cell_value(&mut self.book, sheet, a1, payload)
    }

    pub fn sheet_max_row(&self, sheet: &str) -> PyResult<u32> {
        worksheet::sheet_max_row(&self.book, sheet)
    }

    pub fn sheet_max_column(&self, sheet: &str) -> PyResult<u32> {
        worksheet::sheet_max_column(&self.book, sheet)
    }

    pub fn save(&self, path: &str) -> PyResult<()> {
        let p = Path::new(path);
        writer::xlsx::write(&self.book, p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save: {e}")))
    }
}
