"""Roundtrip tests for freeze panes and split views."""

from __future__ import annotations

from pathlib import Path

import pyumya


def test_freeze_panes_settings_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "freeze.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    ws.freeze_panes = "B2"
    s1 = ws.pane_settings
    assert s1.get("mode") == "freeze"
    assert s1.get("top_left_cell") == "B2"

    ws.set_pane_settings({"mode": "split", "x_split": 1, "y_split": 2})
    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]
    s2 = ws2.pane_settings
    assert s2.get("mode") == "split"
    assert s2.get("x_split") == 1
    assert s2.get("y_split") == 2
