"""Roundtrip tests for data validation rules."""

from __future__ import annotations

from pathlib import Path

import pyumya


def _norm_range(value: object) -> str:
    return str(value or "").replace("$", "").replace("'", "")


def _norm_formula(value: object) -> str:
    if not isinstance(value, str):
        return str(value or "")
    s = value.strip()
    if s.startswith("="):
        s = s[1:]
    if s.startswith('"') and s.endswith('"') and len(s) >= 2:
        s = s[1:-1]
    return s


def _find_validation(
    validations: list[dict[str, object]], *, cell_range: str, formula1: str
) -> dict[str, object] | None:
    for v in validations:
        if _norm_range(v.get("range")) != _norm_range(cell_range):
            continue
        if _norm_formula(v.get("formula1")) != _norm_formula(formula1):
            continue
        return v
    return None


def test_data_validation_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "data_validation.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]

    # 1) List from CSV string
    ws.add_data_validation(
        {
            "range": "B2",
            "validation_type": "list",
            "formula1": '"Red,Green,Blue"',
            "allow_blank": True,
        }
    )

    # 2) List from range
    ws["D2"].value = "A"
    ws["D3"].value = "B"
    ws["D4"].value = "C"
    ws.add_data_validation(
        {
            "range": "B3",
            "validation_type": "list",
            "formula1": "=$D$2:$D$4",
        }
    )

    # 3) Cross-sheet list via named range (we only roundtrip the formula string)
    ws.add_data_validation(
        {
            "range": "B4",
            "validation_type": "list",
            "formula1": "=RegionList",
        }
    )

    # 4) Custom formula
    ws["C5"].value = 5
    ws.add_data_validation(
        {
            "range": "B5",
            "validation_type": "custom",
            "formula1": "=B5>C5",
        }
    )

    # 5) Whole number between with error message
    ws.add_data_validation(
        {
            "range": "B6",
            "validation_type": "whole",
            "operator": "between",
            "formula1": "1",
            "formula2": "10",
            "allow_blank": False,
            "error_title": "Invalid",
            "error": "Enter 1-10",
        }
    )

    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]
    validations = ws2.data_validations

    v1 = _find_validation(validations, cell_range="B2", formula1='"Red,Green,Blue"')
    assert v1 is not None
    assert v1.get("validation_type") == "list"
    assert v1.get("allow_blank") is True

    v2 = _find_validation(validations, cell_range="B3", formula1="=$D$2:$D$4")
    assert v2 is not None
    assert v2.get("validation_type") == "list"

    v3 = _find_validation(validations, cell_range="B4", formula1="=RegionList")
    assert v3 is not None
    assert v3.get("validation_type") == "list"

    v4 = _find_validation(validations, cell_range="B5", formula1="=B5>C5")
    assert v4 is not None
    assert v4.get("validation_type") == "custom"

    v5 = _find_validation(validations, cell_range="B6", formula1="1")
    assert v5 is not None
    assert v5.get("validation_type") == "whole"
    assert v5.get("operator") == "between"
    assert v5.get("formula2") == "10"
    assert v5.get("allow_blank") is False
    assert v5.get("error_title") == "Invalid"
    assert v5.get("error") == "Enter 1-10"
