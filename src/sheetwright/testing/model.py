"""User-facing test model: set/get/recalc over a calc engine."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

from sheetwright.calc import CalcEngine, CalcResult, get_calc_engine
from sheetwright.model.cell import Cell, CellValue
from sheetwright.model.workbook import Workbook
from sheetwright.project import Project
from sheetwright.security import SecurityLimits, get_operator_limits
from sheetwright.source.reader import read_source
from sheetwright.testing.addresses import parse_address
from sheetwright.xlsx.writer import write_xlsx


class Model:
    """A workbook held in memory plus the calc engine that evaluates it.

    `set` mutates the in-memory workbook and invalidates calculated
    values. `get` triggers a recalc on demand and returns the
    calculated value (or the literal value for non-formula cells).
    """

    def __init__(
        self,
        wb: Workbook,
        engine: CalcEngine,
        limits: SecurityLimits = SecurityLimits.defaults(),
    ):
        self._wb = wb
        self._engine = engine
        self._limits = limits
        self._calculated: Optional[CalcResult] = None

    @classmethod
    def open(cls, project_path: str | Path) -> 'Model':
        project = Project.open(project_path)
        wb = read_source(project.root)
        limits = SecurityLimits.effective(
            get_operator_limits(), project.config.security
        )
        return cls(wb, get_calc_engine(project.config.calc_engine), limits)

    @property
    def workbook(self) -> Workbook:
        return self._wb

    def set(self, address: str, value: CellValue) -> None:
        sheet_name, addr = parse_address(self._wb, address)
        sheet = self._wb.sheet(sheet_name)
        existing = sheet.get(addr)
        sheet.set(addr, Cell(value=value, format_id=existing.format_id))
        self._calculated = None

    def get(self, address: str) -> CellValue:
        sheet_name, addr = parse_address(self._wb, address)
        if self._calculated is None:
            self.recalc()
        assert self._calculated is not None  # for mypy
        sheet = self._calculated.get(sheet_name, {})
        if addr in sheet:
            return sheet[addr]
        source_cell = self._wb.sheet(sheet_name).get(addr)
        if source_cell.formula is not None:
            raise RuntimeError(
                f'calc engine omitted formula cell {sheet_name}!{addr}'
            )
        return source_cell.value

    def recalc(self) -> None:
        with tempfile.TemporaryDirectory(prefix='cshs-model-') as td:
            xlsx = Path(td) / 'model.xlsx'
            write_xlsx(self._wb, xlsx)
            self._calculated = self._engine.evaluate(xlsx, limits=self._limits)
