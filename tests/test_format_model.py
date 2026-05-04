from testsweet import test

from claudesheets.model.cell import Cell
from claudesheets.model.format import Border, CellFormat, Fill, Font, Side
from claudesheets.model.workbook import Sheet


@test
def cellformat_default_is_empty():
    fmt = CellFormat()
    assert fmt.font is None
    assert fmt.fill is None
    assert fmt.border is None
    assert fmt.number_format is None


@test
def font_fields():
    f = Font(
        name='Calibri', size=11.0, bold=True, italic=False, color='FF0000'
    )
    assert f.bold and not f.italic


@test
def fill_fields():
    f = Fill(color='FFFF00')
    assert f.color == 'FFFF00'


@test
def border_fields():
    b = Border(left=Side(style='thin', color='000000'))
    assert b.left.style == 'thin'


@test
def sheet_can_associate_format_with_cell():
    sh = Sheet(name='S')
    sh.formats['bold-red'] = CellFormat(
        font=Font(name='Calibri', size=11.0, bold=True, color='FF0000')
    )
    sh.set('A1', Cell(value='hello', format_id='bold-red'))
    assert sh.get('A1').format_id == 'bold-red'
    assert sh.formats['bold-red'].font.bold is True
