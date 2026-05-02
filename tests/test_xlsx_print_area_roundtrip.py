from pathlib import Path

import openpyxl

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx


def _wb_with_print_area(path: Path, area: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'S'
    ws['A1'] = 1
    ws.print_area = area
    wb.save(path)


def test_reader_picks_up_print_area(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    _wb_with_print_area(src, 'A1:E10')
    wb = read_xlsx(src)
    assert wb.sheet('S').print_area == 'A1:E10'


def test_print_area_round_trips_through_writer(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    _wb_with_print_area(src, 'A1:E10')
    out = tmp_path / 'out.xlsx'
    write_xlsx(read_xlsx(src), out)
    assert read_xlsx(out).sheet('S').print_area == 'A1:E10'


def test_print_area_strips_sheet_prefix_on_read(tmp_path: Path):
    # openpyxl sometimes returns 'Sheet1!$A$1:$E$10'; we normalize to
    # 'A1:E10' (sheet implied; dollar signs stripped).
    src = tmp_path / 'in.xlsx'
    _wb_with_print_area(src, "'S'!$A$1:$E$10")
    wb = read_xlsx(src)
    assert wb.sheet('S').print_area == 'A1:E10'
