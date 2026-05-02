"""Excel "ListObject" tables (a.k.a. structured-reference tables).

NOT the legacy "data tables" what-if feature.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Tuple


@dataclass(frozen=True)
class ListTableColumn:
    name: str
    formula: Optional[str] = None
    totals_label: Optional[str] = None
    totals_function: Optional[str] = None  # 'sum', 'average', etc.


@dataclass(frozen=True)
class ListTable:
    name: str
    ref: str  # e.g. 'A1:C10'
    header_row_count: int = 1
    totals_row_count: int = 0
    columns: Tuple[ListTableColumn, ...] = ()
    style: Optional[str] = None  # built-in Excel style name
