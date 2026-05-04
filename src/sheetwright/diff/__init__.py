"""Workbook diff: pure model + computation, no I/O."""

from __future__ import annotations

from sheetwright.diff.compute import diff_workbooks
from sheetwright.diff.model import (
    CellChange,
    FrozenPanesChange,
    NamedRangeChange,
    PrintAreaChange,
    SheetDiff,
    WorkbookDiff,
)

__all__ = [
    'CellChange',
    'FrozenPanesChange',
    'NamedRangeChange',
    'PrintAreaChange',
    'SheetDiff',
    'WorkbookDiff',
    'diff_workbooks',
]
