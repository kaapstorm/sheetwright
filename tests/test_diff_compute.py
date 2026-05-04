from __future__ import annotations

from testsweet import test

from sheetwright.diff import diff_workbooks
from sheetwright.model.cell import Cell
from sheetwright.model.workbook import NamedRange, Sheet, Workbook


def _wb(name: str = 'x') -> Workbook:
    return Workbook(name=name, sheets=[Sheet(name='S')])


@test
def identical_workbooks_have_empty_diff():
    a = _wb()
    b = _wb()
    assert diff_workbooks(a, b).is_empty()


@test
def added_sheet_shows_in_diff():
    a = _wb()
    b = Workbook(name='x', sheets=[Sheet(name='S'), Sheet(name='New')])
    d = diff_workbooks(a, b)
    assert d.sheets_added == ('New',)
    assert d.sheets_removed == ()


@test
def removed_sheet_shows_in_diff():
    a = Workbook(name='x', sheets=[Sheet(name='S'), Sheet(name='Old')])
    b = _wb()
    d = diff_workbooks(a, b)
    assert d.sheets_removed == ('Old',)


@test
def added_cell_shows_in_sheet_diff():
    a = _wb()
    b = _wb()
    b.sheet('S').set('A1', Cell(value=42))
    d = diff_workbooks(a, b)
    assert len(d.sheets_changed) == 1
    sd = d.sheets_changed[0]
    assert sd.name == 'S'
    assert sd.cells_added == ('A1',)


@test
def removed_cell_shows_in_sheet_diff():
    a = _wb()
    a.sheet('S').set('A1', Cell(value=42))
    b = _wb()
    d = diff_workbooks(a, b)
    sd = d.sheets_changed[0]
    assert sd.cells_removed == ('A1',)


@test
def changed_cell_shows_old_and_new_in_sheet_diff():
    a = _wb()
    a.sheet('S').set('A1', Cell(value=1))
    b = _wb()
    b.sheet('S').set('A1', Cell(value=2))
    d = diff_workbooks(a, b)
    sd = d.sheets_changed[0]
    assert len(sd.cells_changed) == 1
    cc = sd.cells_changed[0]
    assert cc.addr == 'A1'
    assert cc.old_value == 1
    assert cc.new_value == 2


@test
def changed_formula_in_diff():
    a = _wb()
    a.sheet('S').set('A1', Cell(formula='=1+1'))
    b = _wb()
    b.sheet('S').set('A1', Cell(formula='=2+2'))
    cc = diff_workbooks(a, b).sheets_changed[0].cells_changed[0]
    assert cc.old_formula == '=1+1'
    assert cc.new_formula == '=2+2'


@test
def named_range_added():
    a = _wb()
    b = _wb()
    b.named_ranges.append(NamedRange(name='gr', ref='S!$A$1'))
    d = diff_workbooks(a, b)
    assert d.named_ranges_added == ('gr',)


@test
def named_range_changed_ref():
    a = _wb()
    a.named_ranges.append(NamedRange(name='gr', ref='S!$A$1'))
    b = _wb()
    b.named_ranges.append(NamedRange(name='gr', ref='S!$B$2'))
    d = diff_workbooks(a, b)
    assert len(d.named_ranges_changed) == 1
    assert d.named_ranges_changed[0].old_ref == 'S!$A$1'
    assert d.named_ranges_changed[0].new_ref == 'S!$B$2'


@test
def frozen_panes_change():
    a = _wb()
    b = _wb()
    b.sheet('S').frozen_panes = 'B2'
    d = diff_workbooks(a, b)
    sd = d.sheets_changed[0]
    assert sd.frozen_panes_change is not None
    assert sd.frozen_panes_change.old is None
    assert sd.frozen_panes_change.new == 'B2'


@test
def print_area_change():
    a = _wb()
    b = _wb()
    b.sheet('S').print_area = 'A1:E10'
    sd = diff_workbooks(a, b).sheets_changed[0]
    assert sd.print_area_change is not None
    assert sd.print_area_change.new == 'A1:E10'


@test
def column_width_change():
    a = _wb()
    b = _wb()
    b.sheet('S').column_widths['A'] = 18.0
    sd = diff_workbooks(a, b).sheets_changed[0]
    assert 'A' in sd.column_widths_changed
    assert sd.column_widths_changed['A'] == (None, 18.0)
