import dataclasses

from testsweet import catch_exceptions, test

from sheetwright.diff.model import (
    CellChange,
    NamedRangeChange,
    SheetDiff,
    WorkbookDiff,
)


@test
def cell_change_is_frozen():
    cc = CellChange(
        sheet='S',
        addr='A1',
        old_value=1,
        new_value=2,
        old_formula=None,
        new_formula=None,
    )
    assert cc.sheet == 'S'
    with catch_exceptions() as excs:
        cc.sheet = 'T'  # type: ignore[misc]
    assert excs and isinstance(excs[0], dataclasses.FrozenInstanceError)


@test
def sheet_diff_defaults_empty():
    sd = SheetDiff(name='S')
    assert sd.cells_added == ()
    assert sd.cells_removed == ()
    assert sd.cells_changed == ()
    assert sd.column_widths_changed == {}
    assert sd.frozen_panes_change is None


@test
def workbook_diff_defaults_empty():
    wd = WorkbookDiff()
    assert wd.sheets_added == ()
    assert wd.sheets_removed == ()
    assert wd.sheets_changed == ()
    assert wd.named_ranges_added == ()
    assert wd.named_ranges_removed == ()
    assert wd.named_ranges_changed == ()


@test
def workbook_diff_is_empty_when_truly_empty():
    wd = WorkbookDiff()
    assert wd.is_empty() is True


@test
def workbook_diff_is_not_empty_when_a_sheet_added():
    wd = WorkbookDiff(sheets_added=('NewSheet',))
    assert wd.is_empty() is False


@test
def workbook_diff_is_empty_even_with_empty_sheet_diffs():
    """A WorkbookDiff containing only empty SheetDiffs should report
    empty. Defends against manual construction sites that don't filter."""
    wd = WorkbookDiff(sheets_changed=(SheetDiff(name='S'),))
    assert wd.is_empty() is True


@test
def named_range_change_holds_old_and_new():
    nrc = NamedRangeChange(
        name='growth_rate',
        old_ref='Inputs!$B$1',
        new_ref='Inputs!$B$2',
    )
    assert nrc.old_ref != nrc.new_ref
