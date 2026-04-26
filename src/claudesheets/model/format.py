"""Cell-formatting model: font, fill, border, number format."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class Font:
    name: Optional[str] = None
    size: Optional[float] = None
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None  # "single", "double", or None
    color: Optional[str] = None  # hex RGB, e.g. "FF0000"


@dataclass(frozen=True)
class Fill:
    color: Optional[str] = None  # hex RGB; solid fill only in Plan 1


@dataclass(frozen=True)
class Side:
    style: Optional[str] = (
        None  # "thin", "medium", "thick", "dashed", "dotted", "double"
    )
    color: Optional[str] = None


@dataclass(frozen=True)
class Border:
    left: Optional[Side] = None
    right: Optional[Side] = None
    top: Optional[Side] = None
    bottom: Optional[Side] = None


@dataclass(frozen=True)
class CellFormat:
    font: Optional[Font] = None
    fill: Optional[Fill] = None
    border: Optional[Border] = None
    number_format: Optional[str] = None  # e.g. "0.00%", "#,##0.00"
