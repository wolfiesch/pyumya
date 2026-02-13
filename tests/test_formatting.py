"""Roundtrip tests for cell formatting."""

from __future__ import annotations

from pathlib import Path

import pyumya


def test_formatting_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "formatting.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    ws["A1"].value = "Hello"
    ws["A2"].value = 123.456

    ws["A1"].font = pyumya.Font(bold=True, color="FF0000")
    ws["A1"].fill = pyumya.PatternFill(fgColor="FFFF00")
    ws["A1"].border = pyumya.Border(
        bottom=pyumya.Side(style="thin"),
        diagonal=pyumya.Side(style="thin"),
        diagonalUp=True,
    )
    ws["A1"].alignment = pyumya.Alignment(horizontal="center", indent=2)
    ws["A2"].number_format = "0.00"

    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]

    assert ws2["A1"].font.bold is True
    assert ws2["A1"].font.color == "FF0000"
    assert ws2["A1"].fill.fgColor == "FFFF00"
    assert ws2["A1"].fill.fill_type == "solid"
    assert ws2["A1"].border.bottom.style == "thin"
    assert ws2["A1"].border.diagonal.style == "thin"
    assert ws2["A1"].border.diagonalUp is True
    assert ws2["A1"].border.diagonalDown is False
    assert ws2["A1"].alignment.horizontal == "center"
    assert ws2["A1"].alignment.indent == 2
    assert ws2["A2"].number_format == "0.00"
