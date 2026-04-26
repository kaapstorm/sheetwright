from datetime import datetime

from claudesheets.model.cell import Cell
from claudesheets.model.workbook import NamedRange, Sheet, Workbook


def test_workbook_is_empty_by_default():
    wb = Workbook(name='m')
    assert wb.sheets == []
    assert wb.named_ranges == []


def test_sheet_get_set_round_trip():
    sh = Sheet(name='S1')
    sh.set('A1', Cell(value=10))
    sh.set('B2', Cell(formula='=A1*2'))
    assert sh.get('A1').value == 10
    assert sh.get('B2').formula == '=A1*2'


def test_sheet_get_unset_returns_blank_cell():
    sh = Sheet(name='S1')
    assert sh.get('Z99') == Cell()


def test_cell_supports_basic_types():
    Cell(value=1)
    Cell(value=1.5)
    Cell(value='x')
    Cell(value=True)
    Cell(value=datetime(2026, 1, 1))
    Cell(value=None)


def test_cell_cannot_have_value_and_formula_together():
    import pytest

    with pytest.raises(ValueError):
        Cell(value=1, formula='=A1')


def test_workbook_lookup_sheet_by_name():
    wb = Workbook(name='m')
    s = Sheet(name='Inputs')
    wb.sheets.append(s)
    assert wb.sheet('Inputs') is s


def test_named_range_workbook_scope():
    nr = NamedRange(name='growth', scope='workbook', ref='Inputs!B5')
    assert nr.scope == 'workbook'


def test_named_range_sheet_scope_requires_sheet_name():
    import pytest

    with pytest.raises(ValueError):
        NamedRange(name='local', scope='sheet', ref='B5')  # missing sheet
    nr = NamedRange(name='local', scope='sheet', sheet='Inputs', ref='B5')
    assert nr.sheet == 'Inputs'
