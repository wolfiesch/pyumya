"""Workbook — top-level Excel file object."""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

from pyumya._rust import RustWorkbook
from pyumya.worksheet import Worksheet


class Workbook:
    """An Excel workbook (.xlsx).

    Create a new workbook::

        wb = Workbook()
        ws = wb.create_sheet("Sheet1")

    Open an existing workbook::

        wb = load_workbook("report.xlsx")
        ws = wb["Sheet1"]
    """

    def __init__(
        self,
        *,
        _rust_book: RustWorkbook | None = None,
        remove_default_sheet: bool = False,
    ) -> None:
        if _rust_book is not None:
            self._rust = _rust_book
        else:
            self._rust = RustWorkbook(remove_default_sheet=remove_default_sheet)

    @property
    def sheetnames(self) -> list[str]:
        """Return list of sheet names in workbook order."""
        return self._rust.sheet_names()

    def create_sheet(self, title: str) -> Worksheet:
        """Create a new worksheet and return it."""
        self._rust.add_sheet(title)
        return Worksheet(self, title)

    def save(self, filename: str | Path) -> None:
        """Save the workbook to disk."""
        self._rust.save(str(filename))

    def __getitem__(self, name: str) -> Worksheet:
        """Get worksheet by name."""
        if name not in self.sheetnames:
            raise KeyError(f"Worksheet '{name}' does not exist.")
        return Worksheet(self, name)

    def __contains__(self, name: str) -> bool:
        return name in self.sheetnames

    def __iter__(self) -> Iterator[Worksheet]:
        return (self[name] for name in self.sheetnames)

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
