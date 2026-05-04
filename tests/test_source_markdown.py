from testsweet import test

from sheetwright.model.cell import Cell
from sheetwright.model.workbook import Sheet
from sheetwright.source.markdown import dump_table, load_table


@test
def round_trip_simple_values():
    sh = Sheet(name='S')
    sh.set('A1', Cell(value='hello'))
    sh.set('B1', Cell(value=42))
    sh.set('A2', Cell(value=3.14))
    sh.set('B2', Cell(value=True))
    text = dump_table(sh)
    sh2 = Sheet(name='S')
    load_table(sh2, text)
    assert sh2.get('A1').value == 'hello'
    assert sh2.get('B1').value == 42
    assert sh2.get('A2').value == 3.14
    assert sh2.get('B2').value is True


@test
def round_trip_formula():
    sh = Sheet(name='S')
    sh.set('A1', Cell(value=10))
    sh.set('B1', Cell(formula='=A1*2'))
    text = dump_table(sh)
    sh2 = Sheet(name='S')
    load_table(sh2, text)
    assert sh2.get('B1').formula == '=A1*2'


@test
def round_trip_blanks_omitted():
    sh = Sheet(name='S')
    sh.set('A1', Cell(value=1))
    sh.set('C5', Cell(value=2))
    text = dump_table(sh)
    sh2 = Sheet(name='S')
    load_table(sh2, text)
    assert sh2.get('A1').value == 1
    assert sh2.get('B1').value is None
    assert sh2.get('C5').value == 2


@test
def dump_includes_header_and_separator():
    sh = Sheet(name='S')
    sh.set('A1', Cell(value=1))
    text = dump_table(sh)
    lines = text.splitlines()
    assert lines[0].startswith('|')
    assert 'A' in lines[0]
    assert set(lines[1].replace('|', '').strip()) <= {'-', ' '}


@test
def load_ignores_blank_rows_and_extra_whitespace():
    sh = Sheet(name='S')
    text = """\
| (cell) | A | B |
| --- | --- | --- |
|  1 |  hi |   |
|    |     |   |
|  3 |     | =A1+1 |
"""
    load_table(sh, text)
    assert sh.get('A1').value == 'hi'
    assert sh.get('A3').value is None
    assert sh.get('B3').formula == '=A1+1'


@test
def string_with_pipe_is_escaped():
    sh = Sheet(name='S')
    sh.set('A1', Cell(value='a | b'))
    text = dump_table(sh)
    sh2 = Sheet(name='S')
    load_table(sh2, text)
    assert sh2.get('A1').value == 'a | b'
