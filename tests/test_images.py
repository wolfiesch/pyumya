"""Roundtrip tests for embedded images."""

from __future__ import annotations

import base64
from pathlib import Path

import pyumya


_PNG_1X1 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/aXo9mQAAAAASUVORK5CYII="
)


def _write_bytes(path: Path, b64: str) -> None:
    path.write_bytes(base64.b64decode(b64))


def _find_image(images: list[dict[str, object]], cell: str) -> dict[str, object] | None:
    for img in images:
        if img.get("cell") == cell:
            return img
    return None


def test_images_roundtrip(tmp_path: Path) -> None:
    png = tmp_path / "sample.png"
    png2 = tmp_path / "sample2.png"
    _write_bytes(png, _PNG_1X1)
    _write_bytes(png2, _PNG_1X1)

    out = tmp_path / "images.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    ws.add_image("B2", str(png))
    ws.add_image("D6", str(png2), offset=(8, 6))
    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]

    img1 = _find_image(ws2.images, "B2")
    assert img1 is not None
    assert img1.get("anchor") == "oneCell"
    assert str(img1.get("path") or "").startswith("/xl/media/")

    img2 = _find_image(ws2.images, "D6")
    assert img2 is not None
    assert img2.get("anchor") == "oneCell"
    assert img2.get("offset") == [8, 6]
    assert str(img2.get("path") or "").startswith("/xl/media/")
