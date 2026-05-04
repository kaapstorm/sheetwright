"""Calc engine plugin layer.

A calc engine evaluates a built `.xlsx` and returns the calculated
values for every cell. The default engine is LibreOffice headless.
"""

from __future__ import annotations

from sheetwright.calc.base import CalcEngine, CalcResult


def get_calc_engine(name: str) -> CalcEngine:
    if name == 'libreoffice':
        from sheetwright.calc.libreoffice import LibreOfficeEngine

        return LibreOfficeEngine()
    raise ValueError(f'unknown calc engine: {name!r}')


__all__ = ['CalcEngine', 'CalcResult', 'get_calc_engine']
