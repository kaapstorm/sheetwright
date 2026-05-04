from testsweet import test

from sheetwright.diff.check import CheckIssue
from sheetwright.diff.compute import diff_workbooks
from sheetwright.mcp.shaping import (
    check_issues_to_dicts,
    diff_to_dict,
)
from sheetwright.model.cell import Cell
from sheetwright.model.workbook import Sheet, Workbook


def _wb() -> Workbook:
    return Workbook(name='x', sheets=[Sheet(name='S')])


@test
def diff_to_dict_for_empty_diff():
    d = diff_workbooks(_wb(), _wb())
    out = diff_to_dict(d)
    assert out['is_empty'] is True
    assert 'no changes' in out['rendered'].lower()
    assert out['structured']['sheets_added'] == []
    assert out['structured']['sheets_removed'] == []


@test
def diff_to_dict_for_added_cell():
    a = _wb()
    b = _wb()
    b.sheet('S').set('A1', Cell(value=42))
    out = diff_to_dict(diff_workbooks(a, b))
    assert out['is_empty'] is False
    assert out['structured']['sheets_changed'][0]['name'] == 'S'
    assert out['structured']['sheets_changed'][0]['cells_added'] == ['A1']


@test
def check_issues_to_dicts_serialises_each_field():
    issues = [
        CheckIssue(
            kind='dangling_sheet_ref',
            detail='unknown sheet',
            location='S!A1',
        ),
        CheckIssue(kind='manifest_sheet_missing', detail='x', location=None),
    ]
    out = check_issues_to_dicts(issues)
    assert out[0] == {
        'kind': 'dangling_sheet_ref',
        'detail': 'unknown sheet',
        'location': 'S!A1',
    }
    assert out[1] == {
        'kind': 'manifest_sheet_missing',
        'detail': 'x',
        'location': None,
    }
