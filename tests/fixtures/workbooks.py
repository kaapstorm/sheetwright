"""Helpers that build small openpyxl workbooks for round-trip tests.

These are not pytest-unmagic fixtures — they are plain helpers callable
from tests. We keep them in one place so test setup stays consistent.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName


def write_simple_xlsx(path: Path) -> None:
    """A two-sheet workbook with values, a formula, and a named range."""
    wb = openpyxl.Workbook()
    s1 = wb.active
    s1.title = 'Inputs'
    s1['A1'] = 'growth_rate'
    s1['B1'] = 0.04
    s1['A2'] = 'base_revenue'
    s1['B2'] = 1_000_000

    s2 = wb.create_sheet('Outputs')
    s2['A1'] = 'revenue_2027'
    s2['B1'] = '=Inputs!B2 * (1 + Inputs!B1)'

    wb.defined_names['growth_rate'] = DefinedName(
        name='growth_rate', attr_text='Inputs!$B$1'
    )

    wb.save(path)
