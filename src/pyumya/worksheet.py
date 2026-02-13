"""Worksheet wrapper for pyumya."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from pyumya.cell import Cell


if TYPE_CHECKING:  # pragma: no cover
    from pyumya.workbook import Workbook


def _column_index_to_letter(col: int) -> str:
    if col <= 0:
        raise ValueError("Column index must be >= 1")
    n = col
    out: list[str] = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def _column_letter_to_index(col: str) -> int:
    s = col.strip().upper()
    if not s or not s.isalpha():
        raise ValueError(f"Invalid column: {col!r}")
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


class Worksheet:
    def __init__(self, workbook: Workbook, title: str) -> None:
        self._workbook = workbook
        self._title = title

    @property
    def title(self) -> str:
        return self._title

    def __getitem__(self, key: str) -> Cell:
        if not isinstance(key, str):
            raise TypeError("Worksheet keys must be A1 strings")
        a1 = key.strip().upper()
        return Cell(self, a1)

    def cell(self, row: int, column: int, value: Any = None) -> Cell:
        a1 = self._a1_from_row_col(row, column)
        c = Cell(self, a1)
        if value is not None:
            c.value = value
        return c

    def append(self, iterable: Any) -> None:
        values = list(iterable)
        row = self.max_row + 1 if self.max_row > 0 else 1
        for idx, v in enumerate(values, start=1):
            self.cell(row=row, column=idx, value=v)

    def iter_rows(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        min_col: int = 1,
        max_col: int | None = None,
    ) -> Iterator[tuple[Cell, ...]]:
        if max_row is None:
            max_row = self.max_row
        if max_col is None:
            max_col = self.max_column
        if max_row <= 0 or max_col <= 0:
            return iter(())

        def gen() -> Iterator[tuple[Cell, ...]]:
            for r in range(min_row, max_row + 1):
                row_cells = tuple(self.cell(r, c) for c in range(min_col, max_col + 1))
                yield row_cells

        return gen()

    def iter_cols(
        self,
        min_row: int = 1,
        max_row: int | None = None,
        min_col: int = 1,
        max_col: int | None = None,
    ) -> Iterator[tuple[Cell, ...]]:
        if max_row is None:
            max_row = self.max_row
        if max_col is None:
            max_col = self.max_column
        if max_row <= 0 or max_col <= 0:
            return iter(())

        def gen() -> Iterator[tuple[Cell, ...]]:
            for c in range(min_col, max_col + 1):
                col_cells = tuple(self.cell(r, c) for r in range(min_row, max_row + 1))
                yield col_cells

        return gen()

    @property
    def max_row(self) -> int:
        return int(self._workbook._rust.sheet_max_row(self._title))

    @property
    def max_column(self) -> int:
        return int(self._workbook._rust.sheet_max_column(self._title))

    # ---------------------------------------------------------------------
    # Internal helpers used by Cell
    # ---------------------------------------------------------------------

    def _a1_from_row_col(self, row: int, col: int) -> str:
        if row <= 0:
            raise ValueError("Row index must be >= 1")
        if col <= 0:
            raise ValueError("Column index must be >= 1")
        return f"{_column_index_to_letter(col)}{row}"

    def _row_col_from_a1(self, a1: str) -> tuple[int, int]:
        s = a1.strip().upper()
        if not s:
            raise ValueError("Empty cell coordinate")
        letters: list[str] = []
        digits: list[str] = []
        for ch in s:
            if ch.isalpha():
                if digits:
                    raise ValueError(f"Invalid A1 coordinate: {a1!r}")
                letters.append(ch)
            elif ch.isdigit():
                digits.append(ch)
            else:
                raise ValueError(f"Invalid A1 coordinate: {a1!r}")
        if not letters or not digits:
            raise ValueError(f"Invalid A1 coordinate: {a1!r}")
        row = int("".join(digits))
        col = _column_letter_to_index("".join(letters))
        if row <= 0 or col <= 0:
            raise ValueError(f"Invalid A1 coordinate: {a1!r}")
        return row, col

    def _rust_read_cell_payload(self, a1: str) -> dict[str, Any]:
        payload = self._workbook._rust.read_cell_value(self._title, a1)
        if isinstance(payload, dict):
            return payload
        return {"type": "string", "value": payload}

    def _rust_write_cell_payload(self, a1: str, payload: dict[str, Any]) -> None:
        self._workbook._rust.write_cell_value(self._title, a1, payload)
