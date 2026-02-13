"""Cell wrapper for pyumya.

This is a small, openpyxl-inspired API layer over the Rust backend.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import TYPE_CHECKING, Any

from pyumya.styles import Alignment, Border, Font, PatternFill, Side, normalize_rgb


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

    # ------------------------------------------------------------------
    # Formatting
    # ------------------------------------------------------------------

    @property
    def font(self) -> Font:
        d = self._ws._rust_read_cell_format(self._coordinate)
        return Font(
            name=str(d.get("font_name", "Calibri")),
            size=float(d.get("font_size", 11.0)),
            bold=bool(d.get("bold", False)),
            italic=bool(d.get("italic", False)),
            underline=str(d.get("underline", "none")),
            strikethrough=bool(d.get("strikethrough", False)),
            color=normalize_rgb(str(d.get("font_color", "000000"))),
        )

    @font.setter
    def font(self, font: Font) -> None:
        self._ws._rust_write_cell_format(
            self._coordinate,
            {
                "bold": bool(font.bold),
                "italic": bool(font.italic),
                "underline": str(font.underline),
                "strikethrough": bool(font.strikethrough),
                "font_name": str(font.name),
                "font_size": float(font.size),
                "font_color": normalize_rgb(font.color),
            },
        )

    @property
    def fill(self) -> PatternFill:
        d = self._ws._rust_read_cell_format(self._coordinate)
        bg = d.get("bg_color")
        fill_type = d.get("fill_type")
        if fill_type is None:
            fill_type = "solid" if bg is not None else "none"
        return PatternFill(fill_type=str(fill_type), fgColor=normalize_rgb(str(bg or "000000")))

    @fill.setter
    def fill(self, fill: PatternFill) -> None:
        if fill.fill_type == "none":
            self._ws._rust_write_cell_format(self._coordinate, {"fill_type": "none"})
            return
        self._ws._rust_write_cell_format(
            self._coordinate,
            {
                "fill_type": str(fill.fill_type),
                "bg_color": normalize_rgb(fill.fgColor),
            },
        )

    @property
    def alignment(self) -> Alignment:
        d = self._ws._rust_read_cell_format(self._coordinate)
        return Alignment(
            horizontal=str(d.get("h_align", "general")),
            vertical=str(d.get("v_align", "bottom")),
            wrap_text=bool(d.get("wrap", False)),
            text_rotation=int(d.get("rotation", 0)),
            indent=int(d.get("indent", 0)),
        )

    @alignment.setter
    def alignment(self, alignment: Alignment) -> None:
        self._ws._rust_write_cell_format(
            self._coordinate,
            {
                "h_align": str(alignment.horizontal),
                "v_align": str(alignment.vertical),
                "wrap": bool(alignment.wrap_text),
                "rotation": int(alignment.text_rotation),
                "indent": int(alignment.indent),
            },
        )

    @property
    def number_format(self) -> str:
        d = self._ws._rust_read_cell_format(self._coordinate)
        return str(d.get("number_format", "General"))

    @number_format.setter
    def number_format(self, fmt: str) -> None:
        self._ws._rust_write_cell_format(self._coordinate, {"number_format": str(fmt)})

    @property
    def border(self) -> Border:
        d = self._ws._rust_read_cell_border(self._coordinate)

        def side(key: str) -> Side:
            raw = d.get(key)
            if not isinstance(raw, dict):
                return Side()
            return Side(
                style=str(raw.get("style", "none")),
                color=normalize_rgb(str(raw.get("color", "000000"))),
            )

        diag_up = side("diagonal_up")
        diag_down = side("diagonal_down")

        diagonal_up_enabled = diag_up.style != "none"
        diagonal_down_enabled = diag_down.style != "none"
        diag = diag_up if diagonal_up_enabled else diag_down

        return Border(
            left=side("left"),
            right=side("right"),
            top=side("top"),
            bottom=side("bottom"),
            diagonal=diag,
            diagonalUp=diagonal_up_enabled,
            diagonalDown=diagonal_down_enabled,
        )

    @border.setter
    def border(self, border: Border) -> None:
        def sd(s: Side) -> dict[str, str]:
            return {"style": str(s.style), "color": normalize_rgb(str(s.color))}

        diag_side = border.diagonal
        diag_on = diag_side.style != "none"
        diag_up = bool(border.diagonalUp) and diag_on
        diag_down = bool(border.diagonalDown) and diag_on

        diag_up_payload = (
            sd(diag_side) if diag_up else {"style": "none", "color": sd(diag_side)["color"]}
        )
        diag_down_payload = (
            sd(diag_side) if diag_down else {"style": "none", "color": sd(diag_side)["color"]}
        )

        self._ws._rust_write_cell_border(
            self._coordinate,
            {
                "left": sd(border.left),
                "right": sd(border.right),
                "top": sd(border.top),
                "bottom": sd(border.bottom),
                "diagonal_up": diag_up_payload,
                "diagonal_down": diag_down_payload,
            },
        )
