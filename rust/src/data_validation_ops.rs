use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use std::str::FromStr;

use umya_spreadsheet::structs::{
    DataValidation, DataValidationOperatorValues, DataValidationValues, DataValidations, EnumTrait,
};
use umya_spreadsheet::Spreadsheet;

fn normalize_validation_type(input: &str) -> String {
    let s = input.trim();
    if s.eq_ignore_ascii_case("textlength") || s.eq_ignore_ascii_case("text_length") {
        return "textLength".to_string();
    }
    s.to_ascii_lowercase()
}

fn normalize_operator(input: &str) -> String {
    let s = input.trim();
    // Keep the exact spellings used by umya-spreadsheet's enums.
    if s.eq_ignore_ascii_case("greater_than") {
        "greaterThan".to_string()
    } else if s.eq_ignore_ascii_case("greater_than_or_equal") {
        "greaterThanOrEqual".to_string()
    } else if s.eq_ignore_ascii_case("less_than") {
        "lessThan".to_string()
    } else if s.eq_ignore_ascii_case("less_than_or_equal") {
        "lessThanOrEqual".to_string()
    } else if s.eq_ignore_ascii_case("not_between") {
        "notBetween".to_string()
    } else if s.eq_ignore_ascii_case("not_equal") {
        "notEqual".to_string()
    } else {
        s.to_string()
    }
}

pub(crate) fn read_data_validations(
    book: &Spreadsheet,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<Py<PyAny>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let list = PyList::empty(py);
    let Some(dvs) = ws.get_data_validations() else {
        return Ok(list.into_any().unbind());
    };

    for dv in dvs.get_data_validation_list() {
        let d = PyDict::new(py);

        let sqref = dv.get_sequence_of_references().get_sqref();
        if !sqref.is_empty() {
            d.set_item("range", sqref)?;
        }

        d.set_item("validation_type", dv.get_type().get_value_string())?;
        d.set_item("operator", dv.get_operator().get_value_string())?;
        d.set_item("allow_blank", *dv.get_allow_blank())?;
        d.set_item("show_input", *dv.get_show_input_message())?;
        d.set_item("show_error", *dv.get_show_error_message())?;

        let f1 = dv.get_formula1();
        if !f1.is_empty() {
            d.set_item("formula1", f1.to_string())?;
        }
        let f2 = dv.get_formula2();
        if !f2.is_empty() {
            d.set_item("formula2", f2.to_string())?;
        }

        let err_title = dv.get_error_title();
        if !err_title.is_empty() {
            d.set_item("error_title", err_title.to_string())?;
        }
        let err = dv.get_error_message();
        if !err.is_empty() {
            d.set_item("error", err.to_string())?;
        }

        let prompt_title = dv.get_prompt_title();
        if !prompt_title.is_empty() {
            d.set_item("prompt_title", prompt_title.to_string())?;
        }
        let prompt = dv.get_prompt();
        if !prompt.is_empty() {
            d.set_item("prompt", prompt.to_string())?;
        }

        list.append(d)?;
    }

    Ok(list.into_any().unbind())
}

pub(crate) fn add_data_validation(
    book: &mut Spreadsheet,
    sheet: &str,
    validation_dict: &Bound<'_, PyAny>,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let dict = validation_dict
        .cast::<PyDict>()
        .map_err(|_| PyErr::new::<PyValueError, _>("validation must be a dict"))?;

    let range_str = dict
        .get_item("range")?
        .ok_or_else(|| PyErr::new::<PyValueError, _>("validation.range is required"))?
        .extract::<String>()?;
    let typ_str = dict
        .get_item("validation_type")?
        .ok_or_else(|| PyErr::new::<PyValueError, _>("validation.validation_type is required"))?
        .extract::<String>()?;

    let typ_norm = normalize_validation_type(&typ_str);
    let typ = DataValidationValues::from_str(&typ_norm).map_err(|_| {
        PyErr::new::<PyValueError, _>(format!("Invalid validation_type: {typ_str}"))
    })?;

    let mut dv = DataValidation::default();
    dv.set_type(typ);
    dv.get_sequence_of_references_mut().set_sqref(range_str);

    if let Some(op) = dict.get_item("operator")? {
        if !op.is_none() {
            let op_str = op.extract::<String>()?;
            let op_norm = normalize_operator(&op_str);
            let op_val = DataValidationOperatorValues::from_str(&op_norm).map_err(|_| {
                PyErr::new::<PyValueError, _>(format!("Invalid operator: {op_str}"))
            })?;
            dv.set_operator(op_val);
        }
    }

    if let Some(f1) = dict.get_item("formula1")? {
        if !f1.is_none() {
            dv.set_formula1(f1.extract::<String>()?);
        }
    }
    if let Some(f2) = dict.get_item("formula2")? {
        if !f2.is_none() {
            dv.set_formula2(f2.extract::<String>()?);
        }
    }

    if let Some(ab) = dict.get_item("allow_blank")? {
        if !ab.is_none() {
            dv.set_allow_blank(ab.extract::<bool>()?);
        }
    }

    if let Some(si) = dict.get_item("show_input")? {
        if !si.is_none() {
            dv.set_show_input_message(si.extract::<bool>()?);
        }
    }
    if let Some(se) = dict.get_item("show_error")? {
        if !se.is_none() {
            dv.set_show_error_message(se.extract::<bool>()?);
        }
    }

    if let Some(pt) = dict.get_item("prompt_title")? {
        if !pt.is_none() {
            dv.set_prompt_title(pt.extract::<String>()?);
        }
    }
    if let Some(p) = dict.get_item("prompt")? {
        if !p.is_none() {
            dv.set_prompt(p.extract::<String>()?);
        }
    }

    if let Some(et) = dict.get_item("error_title")? {
        if !et.is_none() {
            dv.set_error_title(et.extract::<String>()?);
        }
    }
    if let Some(err) = dict.get_item("error")? {
        if !err.is_none() {
            dv.set_error_message(err.extract::<String>()?);
        }
    }

    if ws.get_data_validations().is_none() {
        ws.set_data_validations(DataValidations::default());
    }
    ws.get_data_validations_mut()
        .expect("data validations must exist")
        .add_data_validation_list(dv);

    Ok(())
}
