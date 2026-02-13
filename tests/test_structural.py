"""Roundtrip tests for structural worksheet features."""

from __future__ import annotations

from pathlib import Path

import pyumya


def test_merge_cells_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "merge.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    ws.merge_cells("B1:D1")
    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]
    assert "B1:D1" in ws2.merged_cells.ranges


def test_freeze_panes_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "freeze.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]
    ws.freeze_panes = "A2"
    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]
    assert ws2.freeze_panes == "A2"


def test_row_height_and_column_width_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "dims.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    ws.row_dimensions[1].height = 20
    ws.column_dimensions["A"].width = 15
    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]
    assert ws2.row_dimensions[1].height == 20
    assert ws2.column_dimensions["A"].width == 15
