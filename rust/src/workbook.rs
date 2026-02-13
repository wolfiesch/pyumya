use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;

use std::path::Path;

use umya_spreadsheet::{new_file, reader, writer, Spreadsheet};

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

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        self.book
            .new_sheet(name)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("{e}")))?;
        Ok(())
    }

    pub fn save(&self, path: &str) -> PyResult<()> {
        let p = Path::new(path);
        writer::xlsx::write(&self.book, p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save: {e}")))
    }
}
