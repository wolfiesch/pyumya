"""Roundtrip tests for hyperlinks."""

from __future__ import annotations

from pathlib import Path

import pyumya


def _find_link(links: list[dict[str, object]], cell: str) -> dict[str, object] | None:
    for link in links:
        if link.get("cell") == cell:
            return link
    return None


def test_hyperlink_external_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "links.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    ws.add_hyperlink(
        "B2",
        "https://example.com/docs",
        display="Example Docs",
        tooltip="Go to docs",
        internal=False,
    )
    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]

    link = _find_link(ws2.hyperlinks, "B2")
    assert link is not None
    assert link.get("target") == "https://example.com/docs"
    assert link.get("tooltip") == "Go to docs"
    assert link.get("internal") is False
    assert ws2["B2"].value == "Example Docs"


def test_hyperlink_internal_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "links_internal.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]
    wb.create_sheet("Targets")

    ws.add_hyperlink(
        "B3",
        "'Targets'!A1",
        display="Go Target",
        tooltip="Jump to target",
        internal=True,
    )
    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]

    link = _find_link(ws2.hyperlinks, "B3")
    assert link is not None
    assert link.get("internal") is True
    target = str(link.get("target") or "").lstrip("#").replace("'", "")
    assert target == "Targets!A1"
    assert link.get("tooltip") == "Jump to target"
    assert ws2["B3"].value == "Go Target"
