from datetime import datetime

from claudesheets.model.cell import Cell
from claudesheets.model.workbook import Sheet, Workbook
from claudesheets.snapshot import (
    Snapshot,
    diff_snapshots,
    snapshot_from_calc_result,
)


def _wb_with_formulas(formula_addrs):
    """A workbook where the named addresses are formula cells."""
    sheets_by_name = {}
    for sheet_name, addr in formula_addrs:
        s = sheets_by_name.setdefault(sheet_name, Sheet(name=sheet_name))
        s.set(addr, Cell(formula='=1+1'))
    wb = Workbook(name='x', sheets=list(sheets_by_name.values()))
    return wb


def test_snapshot_includes_only_formula_cells():
    # Source: A1 is literal, B1 is a formula. CalcResult has both.
    wb = _wb_with_formulas([('S1', 'B1')])
    cr = {'S1': {'A1': 'literal', 'B1': 42}}
    snap = snapshot_from_calc_result(cr, wb)
    assert snap.values == {'S1': {'B1': 42}}


def test_snapshot_normalizes_datetime_to_iso_string():
    wb = _wb_with_formulas([('S1', 'A1')])
    cr = {'S1': {'A1': datetime(2024, 3, 15, 12, 0, 0)}}
    snap = snapshot_from_calc_result(cr, wb)
    assert snap.values['S1']['A1'] == '2024-03-15T12:00:00'


def test_snapshot_round_trips_json(tmp_path):
    wb = _wb_with_formulas([('S1', 'A1')])
    snap = snapshot_from_calc_result({'S1': {'A1': 1}}, wb)
    p = tmp_path / 'snap.json'
    snap.write(p)
    loaded = Snapshot.read(p)
    assert loaded == snap


def test_diff_detects_changed_value():
    wb = _wb_with_formulas([('S', 'A1')])
    a = snapshot_from_calc_result({'S': {'A1': 1}}, wb)
    b = snapshot_from_calc_result({'S': {'A1': 2}}, wb)
    diffs = diff_snapshots(a, b)
    assert diffs == [('S', 'A1', 1, 2)]


def test_diff_detects_added_and_removed():
    wb_a = _wb_with_formulas([('S', 'A1')])
    wb_b = _wb_with_formulas([('S', 'A1'), ('S', 'B1')])
    a = snapshot_from_calc_result({'S': {'A1': 1}}, wb_a)
    b = snapshot_from_calc_result({'S': {'A1': 1, 'B1': 2}}, wb_b)
    diffs = diff_snapshots(a, b)
    assert ('S', 'B1', None, 2) in diffs


def test_no_diff_when_equal():
    wb = _wb_with_formulas([('S', 'A1')])
    a = snapshot_from_calc_result({'S': {'A1': 1}}, wb)
    b = snapshot_from_calc_result({'S': {'A1': 1}}, wb)
    assert diff_snapshots(a, b) == []
