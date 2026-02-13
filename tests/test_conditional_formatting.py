"""Roundtrip tests for conditional formatting rules."""

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
    return s


def _find_rule(
    rules: list[dict[str, object]],
    *,
    cell_range: str,
    rule_type: str,
    formula: str | None = None,
    operator: str | None = None,
) -> dict[str, object] | None:
    for r in rules:
        if _norm_range(r.get("range")) != _norm_range(cell_range):
            continue
        if r.get("rule_type") != rule_type:
            continue
        if formula is not None and _norm_formula(r.get("formula")) != _norm_formula(formula):
            continue
        if operator is not None and r.get("operator") != operator:
            continue
        return r
    return None


def test_conditional_formatting_roundtrip(tmp_path: Path) -> None:
    out = tmp_path / "cf.xlsx"

    wb = pyumya.Workbook()
    ws = wb["Sheet1"]
    wb.create_sheet("Ref")
    wb["Ref"]["A1"].value = 10

    for i in range(2, 10):
        ws[f"B{i}"].value = i - 1

    ws.add_conditional_format(
        {
            "range": "B2:B6",
            "rule_type": "cellIs",
            "operator": "greaterThan",
            "formula": "5",
            "format": {"bg_color": "#FFFF00"},
        }
    )
    ws.add_conditional_format(
        {
            "range": "B2:B6",
            "rule_type": "expression",
            "formula": "=Ref!$A$1>5",
            "format": {"bg_color": "#FF00FF"},
        }
    )
    ws.add_conditional_format({"range": "B2:B6", "rule_type": "dataBar"})
    ws.add_conditional_format({"range": "B2:B6", "rule_type": "colorScale"})
    ws.add_conditional_format(
        {
            "range": "B7:B9",
            "rule_type": "cellIs",
            "operator": "lessThan",
            "formula": "3",
            "stop_if_true": True,
            "format": {"bg_color": "#FF0000"},
        }
    )

    wb.save(out)

    wb2 = pyumya.load_workbook(out)
    ws2 = wb2["Sheet1"]
    rules = ws2.conditional_formats

    r1 = _find_rule(
        rules,
        cell_range="B2:B6",
        rule_type="cellIs",
        formula="5",
        operator="greaterThan",
    )
    assert r1 is not None
    fmt1 = r1.get("format")
    assert isinstance(fmt1, dict)
    assert fmt1.get("bg_color") == "#FFFF00"

    r2 = _find_rule(rules, cell_range="B2:B6", rule_type="expression", formula="=Ref!$A$1>5")
    assert r2 is not None
    fmt2 = r2.get("format")
    assert isinstance(fmt2, dict)
    assert fmt2.get("bg_color") == "#FF00FF"

    assert _find_rule(rules, cell_range="B2:B6", rule_type="dataBar") is not None
    assert _find_rule(rules, cell_range="B2:B6", rule_type="colorScale") is not None

    r3 = _find_rule(
        rules,
        cell_range="B7:B9",
        rule_type="cellIs",
        formula="3",
        operator="lessThan",
    )
    assert r3 is not None
    assert r3.get("stop_if_true") is True
    fmt3 = r3.get("format")
    assert isinstance(fmt3, dict)
    assert fmt3.get("bg_color") == "#FF0000"
