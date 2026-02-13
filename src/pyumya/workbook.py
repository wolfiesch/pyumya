"""Workbook — top-level Excel file object."""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

from pyumya._rust import RustWorkbook


class Workbook:
    """An Excel workbook (.xlsx).

    Create a new workbook::

        wb = Workbook()
        ws = wb.create_sheet("Sheet1")

    Open an existing workbook::

        wb = load_workbook("report.xlsx")
        ws = wb["Sheet1"]
    """

    def __init__(self, *, _rust_book: RustWorkbook | None = None) -> None:
        self._rust = _rust_book or RustWorkbook()

    @property
    def sheetnames(self) -> list[str]:
        """Return list of sheet names in workbook order."""
        return self._rust.sheet_names()

    def create_sheet(self, title: str) -> str:
        """Create a new worksheet and return its name."""
        self._rust.add_sheet(title)
        return title

    def save(self, filename: str | Path) -> None:
        """Save the workbook to disk."""
        self._rust.save(str(filename))

    def __getitem__(self, name: str) -> str:
        """Get worksheet by name. Returns name for now (Worksheet class in Phase 1)."""
        if name not in self.sheetnames:
            raise KeyError(f"Worksheet '{name}' does not exist.")
        return name

    def __contains__(self, name: str) -> bool:
        return name in self.sheetnames

    def __iter__(self) -> Iterator[str]:
        return iter(self.sheetnames)

    def __enter__(self) -> Workbook:
        return self

    def __exit__(self, *args: object) -> None:
        pass


def load_workbook(filename: str | Path) -> Workbook:
    """Open an existing Excel workbook (.xlsx).

    Args:
        filename: Path to the .xlsx file.

    Returns:
        A Workbook object.
    """
    rust_book = RustWorkbook.open(str(filename))
    return Workbook(_rust_book=rust_book)
