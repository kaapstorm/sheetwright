"""Document the boundary of CF style fidelity.

CFStyle captures fill colour and a small font subset. Borders,
gradient fills, named numFmts inside CF rules drop on round-trip.
This test asserts the rule *structure* (kind, ranges, operator,
formula) survives even when the visual style is partly dropped.
"""

from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule as XCellIsRule
from openpyxl.styles import Border, Color, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle

from claudesheets.model.conditional import CellIsRule
from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx


def _wb_with_bordered_cf(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'S'
    rule = XCellIsRule(operator='greaterThan', formula=['0'])
    side = Side(style='thin', color='FF000000')
    rule.dxf = DifferentialStyle(
        fill=PatternFill(fill_type='solid', start_color='FF00FF00'),
        border=Border(left=side, right=side, top=side, bottom=side),
    )
    ws.conditional_formatting.add('A1:A10', rule)
    wb.save(path)


def test_cf_with_border_round_trips_structure_but_drops_border(
    tmp_path: Path,
):
    src = tmp_path / 'in.xlsx'
    _wb_with_bordered_cf(src)
    out = tmp_path / 'out.xlsx'
    write_xlsx(read_xlsx(src), out)
    s = read_xlsx(out).sheet('S')

    assert len(s.conditional_formats) == 1
    cf = s.conditional_formats[0]
    assert isinstance(cf, CellIsRule)
    assert cf.operator == 'greaterThan'
    assert cf.formula == ('0',)
    assert 'A1:A10' in cf.ranges

    assert cf.style is not None
    assert cf.style.fill_color == '00FF00FF00' or (
        cf.style.fill_color == 'FF00FF00'
    )


def test_cf_with_font_bold_round_trips_through_xlsx(tmp_path: Path):
    """font_bold IS supposed to survive xlsx round-trip post-fix."""
    src = tmp_path / 'in_font.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'S'
    rule = XCellIsRule(operator='greaterThan', formula=['0'])
    rule.dxf = DifferentialStyle(
        font=Font(b=True, color=Color(rgb='FFFF0000')),
    )
    ws.conditional_formatting.add('A1:A10', rule)
    wb.save(src)

    out = tmp_path / 'out_font.xlsx'
    write_xlsx(read_xlsx(src), out)
    s = read_xlsx(out).sheet('S')

    cf = s.conditional_formats[0]
    assert isinstance(cf, CellIsRule)
    assert cf.style is not None
    assert cf.style.font_bold is True
    assert cf.style.font_color in ('FFFF0000', 'FFFF0000')
