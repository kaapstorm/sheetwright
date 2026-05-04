"""Data model for workbook diffs.

A `WorkbookDiff` captures structural differences between two
`Workbook` instances. It carries no presentation logic — see
`sheetwright.diff.format` for rendering.

All dataclasses are frozen: a diff is a read-only artifact. Variable-
length collections use tuples.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Optional, Tuple

from sheetwright.model.cell import CellValue


@dataclass(frozen=True)
class CellChange:
    sheet: str
    addr: str
    old_value: CellValue = None
    new_value: CellValue = None
    old_formula: Optional[str] = None
    new_formula: Optional[str] = None


@dataclass(frozen=True)
class NamedRangeChange:
    name: str
    old_ref: str
    new_ref: str


@dataclass(frozen=True)
class FrozenPanesChange:
    old: Optional[str]
    new: Optional[str]


@dataclass(frozen=True)
class PrintAreaChange:
    old: Optional[str]
    new: Optional[str]


@dataclass(frozen=True)
class SheetDiff:
    name: str
    cells_added: Tuple[str, ...] = ()
    cells_removed: Tuple[str, ...] = ()
    cells_changed: Tuple[CellChange, ...] = ()
    column_widths_changed: Dict[
        str, Tuple[Optional[float], Optional[float]]
    ] = field(default_factory=dict)
    frozen_panes_change: Optional[FrozenPanesChange] = None
    print_area_change: Optional[PrintAreaChange] = None
    comments_added: Tuple[str, ...] = ()
    comments_removed: Tuple[str, ...] = ()
    comments_changed: Tuple[str, ...] = ()
    conditional_formats_changed: bool = False
    tables_changed: bool = False
    formats_changed: Tuple[str, ...] = ()


@dataclass(frozen=True)
class WorkbookDiff:
    sheets_added: Tuple[str, ...] = ()
    sheets_removed: Tuple[str, ...] = ()
    sheets_changed: Tuple[SheetDiff, ...] = ()
    named_ranges_added: Tuple[str, ...] = ()
    named_ranges_removed: Tuple[str, ...] = ()
    named_ranges_changed: Tuple[NamedRangeChange, ...] = ()

    def is_empty(self) -> bool:
        if (
            self.sheets_added
            or self.sheets_removed
            or self.named_ranges_added
            or self.named_ranges_removed
            or self.named_ranges_changed
        ):
            return False
        # sheets_changed is only meaningful if any contained SheetDiff
        # actually carries differences. The compute layer filters empty
        # diffs out, but a manual construction site might include them.
        return all(_sheet_diff_is_empty(sd) for sd in self.sheets_changed)


def _sheet_diff_is_empty(sd: 'SheetDiff') -> bool:
    return not (
        sd.cells_added
        or sd.cells_removed
        or sd.cells_changed
        or sd.column_widths_changed
        or sd.frozen_panes_change
        or sd.print_area_change
        or sd.comments_added
        or sd.comments_removed
        or sd.comments_changed
        or sd.conditional_formats_changed
        or sd.tables_changed
        or sd.formats_changed
    )
