"""Smoke tests for pyumya workbook operations."""

import tempfile
from pathlib import Path

import pyumya


def test_create_empty_workbook():
    wb = pyumya.Workbook()
    assert wb.sheetnames == ["Sheet1"]


def test_create_workbook_without_default_sheet():
    wb = pyumya.Workbook(remove_default_sheet=True)
    assert wb.sheetnames == []


def test_create_sheet():
    wb = pyumya.Workbook()
    wb.create_sheet("Data")
    assert wb.sheetnames == ["Sheet1", "Data"]
    assert "Data" in wb


def test_roundtrip_save_load(tmp_path: Path):
    # Write
    wb = pyumya.Workbook(remove_default_sheet=True)
    wb.create_sheet("Sheet1")
    wb.create_sheet("Sheet2")
    out = tmp_path / "test.xlsx"
    wb.save(out)

    # Read back
    wb2 = pyumya.load_workbook(out)
    assert wb2.sheetnames == ["Sheet1", "Sheet2"]


def test_context_manager():
    with pyumya.Workbook() as wb:
        wb.create_sheet("Test")
        assert "Test" in wb


def test_getitem_missing_raises():
    wb = pyumya.Workbook()
    try:
        wb["NoSuchSheet"]
        assert False, "Should have raised KeyError"
    except KeyError:
        pass
