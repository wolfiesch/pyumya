"""Roundtrip tests for formulas."""

from __future__ import annotations

from pathlib import Path

import pyumya


def test_formula_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "formulas.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    ws["A1"].value = 1
    ws["A2"].value = 2
    ws["A3"].value = "=SUM(A1:A2)"
    ws["A4"].value = "=A1+A2"

    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]

    assert ws2["A3"].value == "=SUM(A1:A2)"
    assert ws2["A3"].data_type == "f"
    assert ws2["A4"].value == "=A1+A2"
    assert ws2["A4"].data_type == "f"
