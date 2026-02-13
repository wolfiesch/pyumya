"""Roundtrip tests for comments."""

from __future__ import annotations

from pathlib import Path

import pyumya


def _find_comment(comments: list[dict[str, object]], cell: str) -> dict[str, object] | None:
    for c in comments:
        if c.get("cell") == cell:
            return c
    return None


def test_comment_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "comments.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    ws["B2"].value = "Note"
    ws.add_comment("B2", "Legacy note")
    ws.add_comment("B3", "Another note", author="Alice")
    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]

    c1 = _find_comment(ws2.comments, "B2")
    assert c1 is not None
    assert c1.get("text") == "Legacy note"
    assert c1.get("threaded") is False

    c2 = _find_comment(ws2.comments, "B3")
    assert c2 is not None
    assert c2.get("text") == "Another note"
    assert c2.get("author") == "Alice"
