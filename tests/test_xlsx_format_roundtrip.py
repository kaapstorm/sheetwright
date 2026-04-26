from unmagic import fixture, use

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_formatted_xlsx


@fixture
def round_tripped(tmp_path):
    src = tmp_path / 'fmt.xlsx'
    out = tmp_path / 'fmt-out.xlsx'
    write_formatted_xlsx(src)
    wb = read_xlsx(src)
    write_xlsx(wb, out)
    yield read_xlsx(out)


@use(round_tripped)
def test_font_round_trips():
    sheet = round_tripped().sheet('Format')
    cell = sheet.get('A1')
    fmt = sheet.formats[cell.format_id]
    assert fmt.font.bold is True
    assert fmt.font.color in ('FF0000', 'FFFF0000')
    assert fmt.font.name == 'Calibri'


@use(round_tripped)
def test_number_format_round_trips():
    sheet = round_tripped().sheet('Format')
    fmt_b = sheet.formats[sheet.get('B1').format_id]
    assert fmt_b.number_format == '0.00%'
    fmt_c = sheet.formats[sheet.get('C1').format_id]
    assert fmt_c.number_format == '#,##0.00'


@use(round_tripped)
def test_fill_round_trips():
    sheet = round_tripped().sheet('Format')
    fmt = sheet.formats[sheet.get('D1').format_id]
    assert fmt.fill.color in ('FFFF00', 'FFFFFF00')


@use(round_tripped)
def test_border_round_trips():
    sheet = round_tripped().sheet('Format')
    fmt = sheet.formats[sheet.get('E1').format_id]
    assert fmt.border.left.style == 'thin'
    assert fmt.border.right.style == 'thin'
    assert fmt.border.top.style == 'thin'
    assert fmt.border.bottom.style == 'thin'


@use(round_tripped)
def test_column_widths_round_trip():
    sheet = round_tripped().sheet('Format')
    assert sheet.column_widths['A'] == 18
    assert sheet.column_widths['B'] == 10
