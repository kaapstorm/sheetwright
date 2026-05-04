"""In-memory representation of a single spreadsheet cell."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional, Union

CellValue = Union[None, bool, int, float, str, datetime]


@dataclass(frozen=True)
class Cell:
    """A single cell.

    Either `value` is set (a literal cell) or `formula` is set (a formula
    cell), or both are None (a blank cell). They cannot both be set.
    The `format_id` is a reference into the sheet's format table; it
    is set by readers and resolved by writers.
    """

    value: CellValue = None
    formula: Optional[str] = None
    format_id: Optional[str] = None

    def __post_init__(self) -> None:
        if self.value is not None and self.formula is not None:
            raise ValueError('Cell cannot have both value and formula.')

    @property
    def is_blank(self) -> bool:
        return (
            self.value is None
            and self.formula is None
            and self.format_id is None
        )
