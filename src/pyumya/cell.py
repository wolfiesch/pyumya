"""Cell wrapper for pyumya.

This is a small, openpyxl-inspired API layer over the Rust backend.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import TYPE_CHECKING, Any


if TYPE_CHECKING:  # pragma: no cover
    from pyumya.worksheet import Worksheet


def _is_error_token(s: str) -> bool:
    return s == "#N/A" or (s.startswith("#") and s.endswith("!"))


@dataclass
class Cell:
    _ws: Worksheet
    _coordinate: str

    @property
    def coordinate(self) -> str:
        return self._coordinate

    @property
    def row(self) -> int:
        return self._ws._row_col_from_a1(self._coordinate)[0]

    @property
    def column(self) -> int:
        return self._ws._row_col_from_a1(self._coordinate)[1]

    @property
    def data_type(self) -> str:
        payload = self._ws._rust_read_cell_payload(self._coordinate)
        t = payload.get("type", "blank")
        return {
            "string": "s",
            "number": "n",
            "boolean": "b",
            "formula": "f",
            "error": "e",
            "date": "d",
            "datetime": "d",
            "blank": "n",
        }.get(t, "s")

    @property
    def value(self) -> Any:
        payload = self._ws._rust_read_cell_payload(self._coordinate)
        t = payload.get("type", "blank")

        if t == "blank":
            return None

        if t == "string":
            return str(payload.get("value", ""))

        if t == "number":
            v = payload.get("value")
            if isinstance(v, float) and v.is_integer():
                return int(v)
            return v

        if t == "boolean":
            return bool(payload.get("value"))

        if t == "error":
            return str(payload.get("value", ""))

        if t == "formula":
            f = payload.get("formula") or payload.get("value") or ""
            f = str(f)
            return f if f.startswith("=") else f"={f}"

        if t == "date":
            s = str(payload.get("value", ""))
            return date.fromisoformat(s)

        if t == "datetime":
            s = str(payload.get("value", ""))
            return datetime.fromisoformat(s)

        # Fallback: return raw.
        return payload.get("value")

    @value.setter
    def value(self, val: Any) -> None:
        payload: dict[str, Any]

        if val is None:
            payload = {"type": "blank"}
        elif isinstance(val, bool):
            payload = {"type": "boolean", "value": bool(val)}
        elif isinstance(val, (int, float)):
            payload = {"type": "number", "value": float(val)}
        elif isinstance(val, datetime):
            if val.tzinfo is not None:
                raise ValueError("Timezone-aware datetimes are not supported")
            payload = {"type": "datetime", "value": val.isoformat()}
        elif isinstance(val, date):
            payload = {"type": "date", "value": val.isoformat()}
        elif isinstance(val, str):
            if val.startswith("="):
                payload = {"type": "formula", "formula": val}
            elif _is_error_token(val):
                payload = {"type": "error", "value": val}
            else:
                payload = {"type": "string", "value": val}
        else:
            raise TypeError(
                "Cell.value must be one of: str, int, float, bool, None, datetime, date"
            )

        self._ws._rust_write_cell_payload(self._coordinate, payload)
