from pathlib import Path

import openpyxl

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx


def _wb_with_freeze(path: Path, freeze: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'S'
    ws['A1'] = 'header'
    ws.freeze_panes = freeze
    wb.save(path)


def test_xlsx_reader_picks_up_freeze_panes(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    _wb_with_freeze(src, 'B2')
    wb = read_xlsx(src)
    assert wb.sheet('S').frozen_panes == 'B2'


def test_xlsx_writer_emits_freeze_panes(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    _wb_with_freeze(src, 'C5')
    wb = read_xlsx(src)
    out = tmp_path / 'out.xlsx'
    write_xlsx(wb, out)
    re_read = read_xlsx(out)
    assert re_read.sheet('S').frozen_panes == 'C5'


def test_no_freeze_round_trips_as_none(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    wb = openpyxl.Workbook()
    wb.active.title = 'S'
    wb.save(src)
    out = tmp_path / 'out.xlsx'
    write_xlsx(read_xlsx(src), out)
    assert read_xlsx(out).sheet('S').frozen_panes is None
