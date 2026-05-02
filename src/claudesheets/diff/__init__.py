"""Workbook diff: pure model + computation, no I/O."""

from __future__ import annotations

from claudesheets.diff.model import (
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
]
