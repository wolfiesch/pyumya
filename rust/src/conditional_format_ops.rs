use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use std::str::FromStr;

use umya_spreadsheet::structs::{
    Color, ColorScale, ConditionalFormatValueObject, ConditionalFormatValueObjectValues,
    ConditionalFormatValues, ConditionalFormatting, ConditionalFormattingOperatorValues,
    ConditionalFormattingRule, DataBar, EnumTrait, Formula, SequenceOfReferences, Style,
};
use umya_spreadsheet::Spreadsheet;

use crate::utils::{argb_to_hex, hex_to_argb};

pub(crate) fn read_conditional_formats(
    book: &Spreadsheet,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<Py<PyAny>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let list = PyList::empty(py);

    for cf in ws.get_conditional_formatting_collection() {
        let sqref = cf.get_sequence_of_references().get_sqref().to_string();

        for rule in cf.get_conditional_collection() {
            let d = PyDict::new(py);

            if !sqref.is_empty() {
                d.set_item("range", sqref.clone())?;
            }
            d.set_item("rule_type", rule.get_type().get_value_string())?;
            d.set_item("operator", rule.get_operator().get_value_string())?;
            d.set_item("priority", *rule.get_priority())?;
            d.set_item("stop_if_true", *rule.get_stop_if_true())?;

            if let Some(formula) = rule.get_formula() {
                let f = formula.get_address_str();
                if !f.is_empty() {
                    d.set_item("formula", f)?;
                }
            }

            if let Some(style) = rule.get_style() {
                if let Some(color) = style.get_background_color() {
                    let argb = color.get_argb();
                    if !argb.is_empty() {
                        let fmt = PyDict::new(py);
                        fmt.set_item("bg_color", format!("#{}", argb_to_hex(argb)))?;
                        d.set_item("format", fmt)?;
                    }
                }
            }

            list.append(d)?;
        }
    }

    Ok(list.into_any().unbind())
}

fn default_data_bar() -> DataBar {
    let mut bar = DataBar::default();

    let mut min = ConditionalFormatValueObject::default();
    min.set_type(ConditionalFormatValueObjectValues::Number)
        .set_val("0");
    let mut max = ConditionalFormatValueObject::default();
    max.set_type(ConditionalFormatValueObjectValues::Number)
        .set_val("10");
    bar.add_cfvo_collection(min);
    bar.add_cfvo_collection(max);

    let mut c = Color::default();
    c.set_argb(hex_to_argb("638EC6"));
    bar.add_color_collection(c);

    bar
}

fn default_color_scale() -> ColorScale {
    let mut scale = ColorScale::default();

    let mut v1 = ConditionalFormatValueObject::default();
    v1.set_type(ConditionalFormatValueObjectValues::Min);
    let mut v2 = ConditionalFormatValueObject::default();
    v2.set_type(ConditionalFormatValueObjectValues::Percentile)
        .set_val("50");
    let mut v3 = ConditionalFormatValueObject::default();
    v3.set_type(ConditionalFormatValueObjectValues::Max);

    scale.add_cfvo_collection(v1);
    scale.add_cfvo_collection(v2);
    scale.add_cfvo_collection(v3);

    let mut c1 = Color::default();
    c1.set_argb(hex_to_argb("F8696B"));
    let mut c2 = Color::default();
    c2.set_argb(hex_to_argb("FFEB84"));
    let mut c3 = Color::default();
    c3.set_argb(hex_to_argb("63BE7B"));
    scale.add_color_collection(c1);
    scale.add_color_collection(c2);
    scale.add_color_collection(c3);

    scale
}

pub(crate) fn add_conditional_format(
    book: &mut Spreadsheet,
    sheet: &str,
    rule_dict: &Bound<'_, PyAny>,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let dict = rule_dict
        .cast::<PyDict>()
        .map_err(|_| PyErr::new::<PyValueError, _>("cf_rule must be a dict"))?;

    let range_str = dict
        .get_item("range")?
        .ok_or_else(|| PyErr::new::<PyValueError, _>("cf_rule.range is required"))?
        .extract::<String>()?;

    let typ_str = dict
        .get_item("rule_type")?
        .ok_or_else(|| PyErr::new::<PyValueError, _>("cf_rule.rule_type is required"))?
        .extract::<String>()?;

    let typ = ConditionalFormatValues::from_str(&typ_str)
        .map_err(|_| PyErr::new::<PyValueError, _>(format!("Invalid rule_type: {typ_str}")))?;

    let mut rule = ConditionalFormattingRule::default();
    rule.set_type(typ);

    if let Some(op) = dict.get_item("operator")? {
        if !op.is_none() {
            let op_str = op.extract::<String>()?;
            let op_val = ConditionalFormattingOperatorValues::from_str(&op_str).map_err(|_| {
                PyErr::new::<PyValueError, _>(format!("Invalid operator: {op_str}"))
            })?;
            rule.set_operator(op_val);
        }
    }

    if let Some(f) = dict.get_item("formula")? {
        if !f.is_none() {
            let formula_str = f.extract::<String>()?;
            let mut formula = Formula::default();
            formula.set_string_value(formula_str);
            rule.set_formula(formula);
        }
    }

    if let Some(prio) = dict.get_item("priority")? {
        if !prio.is_none() {
            rule.set_priority(prio.extract::<i32>()?);
        }
    }

    if let Some(sit) = dict.get_item("stop_if_true")? {
        if !sit.is_none() {
            rule.set_stop_if_true(sit.extract::<bool>()?);
        }
    }

    if let Some(fmt) = dict.get_item("format")? {
        if !fmt.is_none() {
            let fmt_dict = fmt
                .cast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("cf_rule.format must be a dict"))?;
            if let Some(bg) = fmt_dict.get_item("bg_color")? {
                if !bg.is_none() {
                    let bg_str = bg.extract::<String>()?;
                    let argb = hex_to_argb(&bg_str);
                    let mut style = Style::default();
                    style.set_background_color(argb);
                    rule.set_style(style);
                }
            }
        }
    }

    // Ensure rule contents are present for CF types that require an inner element.
    match rule.get_type() {
        ConditionalFormatValues::DataBar => {
            if rule.get_data_bar().is_none() {
                rule.set_data_bar(default_data_bar());
            }
        }
        ConditionalFormatValues::ColorScale => {
            if rule.get_color_scale().is_none() {
                rule.set_color_scale(default_color_scale());
            }
        }
        _ => {}
    }

    let mut seq = SequenceOfReferences::default();
    seq.set_sqref(range_str);
    let mut cf = ConditionalFormatting::default();
    cf.set_sequence_of_references(seq);
    cf.set_conditional_collection(vec![rule]);

    ws.add_conditional_formatting_collection(cf);
    Ok(())
}
