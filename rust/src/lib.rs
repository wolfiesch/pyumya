use pyo3::prelude::*;

mod workbook;

/// pyumya._rust -- Rust backend for pyumya
///
/// This module provides the low-level Rust bindings to umya-spreadsheet.
/// The Python-facing API lives in the `pyumya` package; this module is internal.
#[pymodule]
fn _rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<workbook::RustWorkbook>()?;
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}
