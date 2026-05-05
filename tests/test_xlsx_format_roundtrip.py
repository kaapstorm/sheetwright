import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from sheetwright.xlsx.reader import read_xlsx
from sheetwright.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_formatted_xlsx
from sheetwright.security import SecurityLimits


@contextmanager
def _round_tripped():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'fmt.xlsx'
        out = tmp_path / 'fmt-out.xlsx'
        write_formatted_xlsx(src)
        wb = read_xlsx(src, limits=SecurityLimits.defaults())
        write_xlsx(wb, out)
        yield read_xlsx(out, limits=SecurityLimits.defaults())


@test
def font_round_trips():
    with _round_tripped() as wb:
        sheet = wb.sheet('Format')
        cell = sheet.get('A1')
        fmt = sheet.formats[cell.format_id]
        assert fmt.font.bold is True
        assert fmt.font.color in ('FF0000', 'FFFF0000')
        assert fmt.font.name == 'Calibri'


@test
def number_format_round_trips():
    with _round_tripped() as wb:
        sheet = wb.sheet('Format')
        fmt_b = sheet.formats[sheet.get('B1').format_id]
        assert fmt_b.number_format == '0.00%'
        fmt_c = sheet.formats[sheet.get('C1').format_id]
        assert fmt_c.number_format == '#,##0.00'


@test
def fill_round_trips():
    with _round_tripped() as wb:
        sheet = wb.sheet('Format')
        fmt = sheet.formats[sheet.get('D1').format_id]
        assert fmt.fill.color in ('FFFF00', 'FFFFFF00')


@test
def border_round_trips():
    with _round_tripped() as wb:
        sheet = wb.sheet('Format')
        fmt = sheet.formats[sheet.get('E1').format_id]
        assert fmt.border.left.style == 'thin'
        assert fmt.border.right.style == 'thin'
        assert fmt.border.top.style == 'thin'
        assert fmt.border.bottom.style == 'thin'


@test
def column_widths_round_trip():
    with _round_tripped() as wb:
        sheet = wb.sheet('Format')
        assert sheet.column_widths['A'] == 18
        assert sheet.column_widths['B'] == 10
