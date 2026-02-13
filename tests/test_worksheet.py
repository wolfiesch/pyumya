"""Worksheet API tests."""

from __future__ import annotations

import pyumya


def test_worksheet_cell_access() -> None:
    wb = pyumya.Workbook()
    ws = wb.create_sheet("Sheet1")

    c = ws["A1"]
    assert c.coordinate == "A1"
    assert c.row == 1
    assert c.column == 1
    assert c.value is None

    c2 = ws.cell(row=2, column=3, value="X")
    assert c2.coordinate == "C2"
    assert ws["C2"].value == "X"


def test_iter_rows_cols_and_append() -> None:
    wb = pyumya.Workbook()
    ws = wb.create_sheet("Sheet1")

    assert list(ws.iter_rows()) == []
    assert list(ws.iter_cols()) == []

    ws.append([1, 2, 3])
    assert ws.max_row == 1
    assert ws.max_column == 3

    rows = list(ws.iter_rows())
    assert len(rows) == 1
    assert [c.value for c in rows[0]] == [1, 2, 3]

    cols = list(ws.iter_cols())
    assert len(cols) == 3
    assert [c.value for c in cols[0]] == [1]
    assert [c.value for c in cols[1]] == [2]
    assert [c.value for c in cols[2]] == [3]

    ws.cell(row=2, column=3, value="Y")
    assert ws.max_row == 2
    assert ws.max_column == 3
