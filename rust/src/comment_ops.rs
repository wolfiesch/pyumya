use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::{Comment, RichText};
use umya_spreadsheet::Spreadsheet;

fn extract_comment_text(comment: &Comment) -> String {
    let ct = comment.get_text();

    if let Some(rt) = ct.get_rich_text() {
        return rt.get_text().to_string();
    }

    // Fallback: CommentText may store a plain `Text` node which umya-spreadsheet
    // does not expose as a public string API. Parse the Debug representation.
    let dbg = format!("{:?}", ct);
    let needle = "value: \"";
    let Some(start) = dbg.find(needle) else {
        return String::new();
    };
    let rest = &dbg[start + needle.len()..];
    let Some(end) = rest.find('"') else {
        return String::new();
    };
    rest[..end].to_string()
}

pub(crate) fn read_comments(
    book: &Spreadsheet,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<Py<PyAny>> {
    let ws = book
        .get_sheet_by_name(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let list = PyList::empty(py);

    for comment in ws.get_comments() {
        let d = PyDict::new(py);
        d.set_item("cell", comment.get_coordinate().to_string())?;
        d.set_item("text", extract_comment_text(comment))?;
        let author = comment.get_author();
        if !author.is_empty() {
            d.set_item("author", author.to_string())?;
        }
        d.set_item("threaded", false)?;
        list.append(d)?;
    }

    Ok(list.into_any().unbind())
}

pub(crate) fn add_comment(
    book: &mut Spreadsheet,
    sheet: &str,
    a1: &str,
    text: &str,
    author: Option<&str>,
) -> PyResult<()> {
    let ws = book
        .get_sheet_by_name_mut(sheet)
        .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

    let mut comment = Comment::default();
    comment.new_comment(a1);
    if let Some(a) = author {
        if !a.is_empty() {
            comment.set_author(a);
        }
    }

    let mut rt = RichText::default();
    rt.set_text(text);
    comment.get_text_mut().set_rich_text(rt);

    ws.add_comments(comment);
    Ok(())
}
