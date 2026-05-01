"""CalcEngine ABC and the registry."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict

from claudesheets.model.cell import CellValue

CalcResult = Dict[str, Dict[str, CellValue]]


class CalcEngine(ABC):
    """Abstract calc engine: evaluate a built .xlsx and return values."""

    @abstractmethod
    def evaluate(self, xlsx_path: Path) -> CalcResult:
        """Return calculated values keyed by sheet, then by A1 address."""
