"""In-memory workbook, sheets, and named ranges."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from sheetwright.model.cell import Cell
from sheetwright.model.comment import Comment
from sheetwright.model.conditional import ConditionalFormat
from sheetwright.model.format import CellFormat
from sheetwright.model.table import ListTable
from sheetwright.model.validation import DataValidation


@dataclass
class NamedRange:
    name: str
    ref: str
    scope: str = 'workbook'  # "workbook" or "sheet"
    sheet: Optional[str] = None  # required when scope == "sheet"

    def __post_init__(self) -> None:
        if self.scope not in ('workbook', 'sheet'):
            raise ValueError(f'Invalid scope: {self.scope!r}')
        if self.scope == 'sheet' and not self.sheet:
            raise ValueError('Sheet-scoped named range requires `sheet`.')


@dataclass
class Sheet:
    name: str
    cells: Dict[str, Cell] = field(default_factory=dict)
    column_widths: Dict[str, float] = field(default_factory=dict)
    formats: Dict[str, CellFormat] = field(default_factory=dict)
    validations: List[DataValidation] = field(default_factory=list)
    frozen_panes: Optional[str] = None  # e.g. "B2"; reserved for Plan 3
    print_area: Optional[str] = None
    comments: Dict[str, Comment] = field(default_factory=dict)
    conditional_formats: List[ConditionalFormat] = field(default_factory=list)
    tables: List[ListTable] = field(default_factory=list)

    def get(self, address: str) -> Cell:
        return self.cells.get(address, Cell())

    def set(self, address: str, cell: Cell) -> None:
        if cell.is_blank:
            self.cells.pop(address, None)
        else:
            self.cells[address] = cell

    def addresses(self) -> List[str]:
        return list(self.cells.keys())


@dataclass
class Workbook:
    name: str
    sheets: List[Sheet] = field(default_factory=list)
    named_ranges: List[NamedRange] = field(default_factory=list)

    def sheet(self, name: str) -> Sheet:
        for s in self.sheets:
            if s.name == name:
                return s
        raise KeyError(name)
