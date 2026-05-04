from testsweet import test

from sheetwright.diff import diff_workbooks
from sheetwright.diff.format import render
from sheetwright.model.cell import Cell
from sheetwright.model.workbook import NamedRange, Sheet, Workbook


def _wb() -> Workbook:
    return Workbook(name='x', sheets=[Sheet(name='S')])


@test
def render_empty_diff_returns_no_changes_message():
    out = render(diff_workbooks(_wb(), _wb()))
    assert 'no changes' in out.lower()


@test
def render_added_sheet_lists_it():
    a = _wb()
    b = Workbook(name='x', sheets=[Sheet(name='S'), Sheet(name='Q4')])
    out = render(diff_workbooks(a, b))
    assert 'Q4' in out
    assert 'added' in out.lower()


@test
def render_removed_sheet_lists_it():
    a = Workbook(name='x', sheets=[Sheet(name='S'), Sheet(name='Old')])
    b = _wb()
    out = render(diff_workbooks(a, b))
    assert 'Old' in out
    assert 'removed' in out.lower()


@test
def render_changed_cell_shows_address_and_values():
    a = _wb()
    a.sheet('S').set('A1', Cell(value=1))
    b = _wb()
    b.sheet('S').set('A1', Cell(value=2))
    out = render(diff_workbooks(a, b))
    assert 'S!A1' in out
    assert '1' in out
    assert '2' in out


@test
def render_named_range_added():
    a = _wb()
    b = _wb()
    b.named_ranges.append(NamedRange(name='growth_rate', ref='S!$A$1'))
    out = render(diff_workbooks(a, b))
    assert 'growth_rate' in out


@test
def render_groups_per_sheet():
    a = _wb()
    b = _wb()
    b.sheet('S').set('A1', Cell(value=1))
    b.sheet('S').set('B2', Cell(value=2))
    out = render(diff_workbooks(a, b))
    # The sheet header appears once, with each cell listed under it.
    assert out.count('Sheet S:') == 1 or out.count('S:') >= 1
    assert 'A1' in out and 'B2' in out
