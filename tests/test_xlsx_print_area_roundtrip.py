import tempfile
from contextlib import contextmanager
from pathlib import Path

import openpyxl
from testsweet import test

from sheetwright.xlsx.reader import read_xlsx
from sheetwright.xlsx.writer import write_xlsx
from sheetwright.security import SecurityLimits


@contextmanager
def _tmp_path():
    with tempfile.TemporaryDirectory() as td:
        yield Path(td)


def _wb_with_print_area(path: Path, area: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'S'
    ws['A1'] = 1
    ws.print_area = area
    wb.save(path)


@test
def reader_picks_up_print_area():
    with _tmp_path() as tmp_path:
        src = tmp_path / 'in.xlsx'
        _wb_with_print_area(src, 'A1:E10')
        wb = read_xlsx(src, limits=SecurityLimits.defaults())
        assert wb.sheet('S').print_area == 'A1:E10'


@test
def print_area_round_trips_through_writer():
    with _tmp_path() as tmp_path:
        src = tmp_path / 'in.xlsx'
        _wb_with_print_area(src, 'A1:E10')
        out = tmp_path / 'out.xlsx'
        write_xlsx(read_xlsx(src, limits=SecurityLimits.defaults()), out)
        assert (
            read_xlsx(out, limits=SecurityLimits.defaults())
            .sheet('S')
            .print_area
            == 'A1:E10'
        )


@test
def print_area_strips_sheet_prefix_on_read():
    # openpyxl sometimes returns 'Sheet1!$A$1:$E$10'; we normalize to
    # 'A1:E10' (sheet implied; dollar signs stripped).
    with _tmp_path() as tmp_path:
        src = tmp_path / 'in.xlsx'
        _wb_with_print_area(src, "'S'!$A$1:$E$10")
        wb = read_xlsx(src, limits=SecurityLimits.defaults())
        assert wb.sheet('S').print_area == 'A1:E10'
