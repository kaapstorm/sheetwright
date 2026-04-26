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


def write_formatted_xlsx(path: Path) -> None:
    """A workbook exercising fonts, fills, borders, number formats, column widths."""
    from openpyxl.styles import (
        Border as XBorder,
        Font as XFont,
        PatternFill,
        Side as XSide,
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Format'

    ws['A1'] = 'bold red'
    ws['A1'].font = XFont(name='Calibri', size=11, bold=True, color='FFFF0000')

    ws['B1'] = 0.1234
    ws['B1'].number_format = '0.00%'

    ws['C1'] = 1234.5
    ws['C1'].number_format = '#,##0.00'

    ws['D1'] = 'filled'
    ws['D1'].fill = PatternFill(fill_type='solid', fgColor='FFFFFF00')

    ws['E1'] = 'boxed'
    side = XSide(style='thin', color='FF000000')
    ws['E1'].border = XBorder(left=side, right=side, top=side, bottom=side)

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 10

    wb.save(path)
