use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;

use std::path::Path;

use umya_spreadsheet::{new_file, reader, writer, Spreadsheet};

use crate::{
    cell_ops, comment_ops, conditional_format_ops, data_validation_ops, format_ops, hyperlink_ops,
    image_ops, structural_ops, worksheet,
};

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
    #[pyo3(signature = (remove_default_sheet = false))]
    pub fn new(remove_default_sheet: bool) -> Self {
        let mut book = new_file();

        if remove_default_sheet {
            // umya-spreadsheet's new_file() includes a default worksheet.
            // Remove it if requested so callers can start from an empty workbook.
            let _ = book.remove_sheet_by_name("Sheet1");
            if !book.get_sheet_collection().is_empty() {
                let _ = book.remove_sheet(0);
            }
        }
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

    // =========================================================================
    // Phase 2: Formatting
    // =========================================================================

    pub fn read_cell_format(&self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<Py<PyAny>> {
        format_ops::read_cell_format(&self.book, py, sheet, a1)
    }

    pub fn write_cell_format(
        &mut self,
        sheet: &str,
        a1: &str,
        format_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        format_ops::write_cell_format(&mut self.book, sheet, a1, format_dict)
    }

    pub fn read_cell_border(&self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<Py<PyAny>> {
        format_ops::read_cell_border(&self.book, py, sheet, a1)
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        border_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        format_ops::write_cell_border(&mut self.book, sheet, a1, border_dict)
    }

    // =========================================================================
    // Phase 2: Structural
    // =========================================================================

    pub fn read_row_height(&self, sheet: &str, row: u32) -> PyResult<Option<f64>> {
        structural_ops::read_row_height(&self.book, sheet, row)
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        structural_ops::set_row_height(&mut self.book, sheet, row, height)
    }

    pub fn read_column_width(&self, sheet: &str, col_str: &str) -> PyResult<Option<f64>> {
        structural_ops::read_column_width(&self.book, sheet, col_str)
    }

    pub fn set_column_width(&mut self, sheet: &str, col_str: &str, width: f64) -> PyResult<()> {
        structural_ops::set_column_width(&mut self.book, sheet, col_str, width)
    }

    pub fn merge_cells(&mut self, sheet: &str, range_str: &str) -> PyResult<()> {
        structural_ops::merge_cells(&mut self.book, sheet, range_str)
    }

    pub fn unmerge_cells(&mut self, sheet: &str, range_str: &str) -> PyResult<()> {
        structural_ops::unmerge_cells(&mut self.book, sheet, range_str)
    }

    pub fn get_merged_ranges(&self, sheet: &str) -> PyResult<Vec<String>> {
        structural_ops::get_merged_ranges(&self.book, sheet)
    }

    pub fn get_freeze_panes(&self, sheet: &str) -> PyResult<Option<String>> {
        structural_ops::get_freeze_panes(&self.book, sheet)
    }

    pub fn set_freeze_panes(&mut self, sheet: &str, a1: Option<&str>) -> PyResult<()> {
        structural_ops::set_freeze_panes(&mut self.book, sheet, a1)
    }

    pub fn read_freeze_panes_settings(&self, py: Python<'_>, sheet: &str) -> PyResult<Py<PyAny>> {
        structural_ops::read_freeze_panes_settings(&self.book, py, sheet)
    }

    pub fn set_freeze_panes_settings(
        &mut self,
        sheet: &str,
        settings: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        structural_ops::set_freeze_panes_settings(&mut self.book, sheet, settings)
    }

    // =========================================================================
    // Tier 2: Hyperlinks
    // =========================================================================

    pub fn read_hyperlinks(&self, py: Python<'_>, sheet: &str) -> PyResult<Py<PyAny>> {
        hyperlink_ops::read_hyperlinks(&self.book, py, sheet)
    }

    pub fn add_hyperlink(
        &mut self,
        sheet: &str,
        a1: &str,
        target: &str,
        tooltip: Option<&str>,
        internal: bool,
    ) -> PyResult<()> {
        hyperlink_ops::add_hyperlink(&mut self.book, sheet, a1, target, tooltip, internal)
    }

    // =========================================================================
    // Tier 2: Comments
    // =========================================================================

    pub fn read_comments(&self, py: Python<'_>, sheet: &str) -> PyResult<Py<PyAny>> {
        comment_ops::read_comments(&self.book, py, sheet)
    }

    pub fn add_comment(
        &mut self,
        sheet: &str,
        a1: &str,
        text: &str,
        author: Option<&str>,
    ) -> PyResult<()> {
        comment_ops::add_comment(&mut self.book, sheet, a1, text, author)
    }

    // =========================================================================
    // Tier 2: Data Validation
    // =========================================================================

    pub fn read_data_validations(&self, py: Python<'_>, sheet: &str) -> PyResult<Py<PyAny>> {
        data_validation_ops::read_data_validations(&self.book, py, sheet)
    }

    pub fn add_data_validation(
        &mut self,
        sheet: &str,
        validation_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        data_validation_ops::add_data_validation(&mut self.book, sheet, validation_dict)
    }

    // =========================================================================
    // Tier 2: Conditional Formatting
    // =========================================================================

    pub fn read_conditional_formats(&self, py: Python<'_>, sheet: &str) -> PyResult<Py<PyAny>> {
        conditional_format_ops::read_conditional_formats(&self.book, py, sheet)
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        rule_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        conditional_format_ops::add_conditional_format(&mut self.book, sheet, rule_dict)
    }

    // =========================================================================
    // Tier 2: Images
    // =========================================================================

    pub fn read_images(&self, py: Python<'_>, sheet: &str) -> PyResult<Py<PyAny>> {
        image_ops::read_images(&self.book, py, sheet)
    }

    pub fn add_image(
        &mut self,
        sheet: &str,
        cell: &str,
        path: &str,
        offset: Option<(i32, i32)>,
    ) -> PyResult<()> {
        image_ops::add_image(&mut self.book, sheet, cell, path, offset)
    }

    pub fn save(&self, path: &str) -> PyResult<()> {
        let p = Path::new(path);
        writer::xlsx::write(&self.book, p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save: {e}")))
    }
}
