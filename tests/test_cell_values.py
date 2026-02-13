"""Roundtrip tests for core cell value types."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

import pyumya


def test_cell_value_roundtrip_all_types(tmp_path: Path) -> None:
    out = tmp_path / "values.xlsx"

    wb = pyumya.Workbook()
    ws = wb.create_sheet("Sheet1")

    ws["A1"].value = "Hello"
    ws["A2"].value = 42
    ws["A3"].value = 3.14
    ws["A4"].value = True
    ws["A5"].value = None
    ws["A6"].value = date(2026, 2, 1)
    ws["A7"].value = datetime(2026, 2, 1, 12, 34, 56)
    ws["A8"].value = "#DIV/0!"
    ws["A9"].value = "#N/A"
    ws["A10"].value = "#VALUE!"

    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]

    assert ws2["A1"].value == "Hello"
    assert ws2["A2"].value == 42
    assert ws2["A3"].value == pytest.approx(3.14)
    assert ws2["A4"].value is True
    assert ws2["A5"].value is None
    assert ws2["A6"].value == date(2026, 2, 1)
    assert ws2["A7"].value == datetime(2026, 2, 1, 12, 34, 56)
    assert ws2["A8"].value == "#DIV/0!"
    assert ws2["A9"].value == "#N/A"
    assert ws2["A10"].value == "#VALUE!"
