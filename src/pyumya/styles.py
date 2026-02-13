"""Style objects for pyumya.

These are intentionally small, openpyxl-inspired value objects.
"""

from __future__ import annotations

from dataclasses import dataclass, field


def normalize_rgb(color: str) -> str:
    """Normalize a color string to 6-digit uppercase RGB (RRGGBB).

    Accepts: "RRGGBB", "#RRGGBB", "FFRRGGBB".
    """

    s = color.strip()
    if s.startswith("#"):
        s = s[1:]
    s = s.upper()

    if len(s) == 8:
        # Drop alpha.
        s = s[2:]
    if len(s) != 6:
        raise ValueError(f"Invalid color string: {color!r}")
    return s


@dataclass
class Font:
    name: str = "Calibri"
    size: float = 11.0
    bold: bool = False
    italic: bool = False
    underline: str = "none"
    strikethrough: bool = False
    color: str = "000000"

    def __post_init__(self) -> None:
        self.color = normalize_rgb(self.color)


@dataclass
class PatternFill:
    fill_type: str = "none"
    fgColor: str = "000000"

    def __post_init__(self) -> None:
        self.fgColor = normalize_rgb(self.fgColor)
        # Convenience: if a non-default color is provided, assume solid.
        if self.fill_type == "none" and self.fgColor != "000000":
            self.fill_type = "solid"


@dataclass
class Side:
    style: str = "none"
    color: str = "000000"

    def __post_init__(self) -> None:
        self.color = normalize_rgb(self.color)


@dataclass
class Border:
    left: Side = field(default_factory=Side)
    right: Side = field(default_factory=Side)
    top: Side = field(default_factory=Side)
    bottom: Side = field(default_factory=Side)
    diagonal: Side = field(default_factory=Side)


@dataclass
class Alignment:
    horizontal: str = "general"
    vertical: str = "bottom"
    wrap_text: bool = False
    text_rotation: int = 0
