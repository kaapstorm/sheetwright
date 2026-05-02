from pathlib import Path

import openpyxl
from openpyxl.comments import Comment as XComment

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx


def _wb_with_comment(path: Path, addr: str, author: str, text: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'S'
    ws[addr] = 1
    ws[addr].comment = XComment(text, author)
    wb.save(path)


def test_reader_picks_up_comment(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    _wb_with_comment(src, 'B2', 'Alice', 'Check this')
    wb = read_xlsx(src)
    s = wb.sheet('S')
    assert 'B2' in s.comments
    assert s.comments['B2'].author == 'Alice'
    assert s.comments['B2'].text == 'Check this'


def test_comment_round_trips_through_writer(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    _wb_with_comment(src, 'B2', 'Alice', 'Check this')
    out = tmp_path / 'out.xlsx'
    write_xlsx(read_xlsx(src), out)
    s = read_xlsx(out).sheet('S')
    assert s.comments['B2'].text == 'Check this'
    assert s.comments['B2'].author == 'Alice'
