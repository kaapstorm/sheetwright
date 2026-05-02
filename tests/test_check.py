from pathlib import Path

from claudesheets.diff.check import (
    CheckIssue,
    check_workbook,
)
from claudesheets.model.cell import Cell
from claudesheets.model.workbook import NamedRange, Sheet, Workbook
from claudesheets.project import Project


def _project(tmp_path: Path, sheets: list[str]) -> Project:
    p = tmp_path / 'proj'
    p.mkdir(exist_ok=True)
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    sheets_toml = '\n'.join(f'    "{s}",' for s in sheets)
    (p / 'workbook.toml').write_text(
        f'[workbook]\nname = "x"\nsheets = [\n{sheets_toml}\n]\n'
    )
    (p / 'sheets').mkdir(exist_ok=True)
    (p / 'data').mkdir(exist_ok=True)
    return Project.open(p)


def test_clean_workbook_has_no_issues(tmp_path: Path):
    project = _project(tmp_path, ['01_inputs'])
    (project.sheets_dir / '01_inputs.md').write_text(
        '| (cell) | A |\n| --- | --- |\n'
    )
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    issues = check_workbook(wb, project)
    assert issues == []


def test_dangling_sheet_ref_reported(tmp_path: Path):
    project = _project(tmp_path, [])
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    wb.sheet('Inputs').set('A1', Cell(formula='=Outputs!B1'))
    issues = check_workbook(wb, project)
    kinds = {i.kind for i in issues}
    assert 'dangling_sheet_ref' in kinds


def test_dangling_named_range_reported(tmp_path: Path):
    project = _project(tmp_path, [])
    wb = Workbook(name='x', sheets=[Sheet(name='S')])
    wb.sheet('S').set('A1', Cell(formula='=growth_rate*2'))
    issues = check_workbook(wb, project)
    kinds = {i.kind for i in issues}
    assert 'dangling_named_range' in kinds


def test_named_reference_resolved_when_defined(tmp_path: Path):
    project = _project(tmp_path, [])
    wb = Workbook(name='x', sheets=[Sheet(name='S')])
    wb.sheet('S').set('A1', Cell(formula='=growth_rate*2'))
    wb.named_ranges.append(NamedRange(name='growth_rate', ref='S!$B$1'))
    issues = check_workbook(wb, project)
    kinds = {i.kind for i in issues}
    assert 'dangling_named_range' not in kinds


def test_manifest_sheet_missing(tmp_path: Path):
    project = _project(tmp_path, ['01_missing'])
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    issues = check_workbook(wb, project)
    kinds = {i.kind for i in issues}
    assert 'manifest_sheet_missing' in kinds


def test_sheet_file_missing_from_manifest(tmp_path: Path):
    project = _project(tmp_path, [])
    (project.sheets_dir / '01_orphan.md').write_text(
        '| (cell) | A |\n| --- | --- |\n'
    )
    wb = Workbook(name='x', sheets=[Sheet(name='Orphan')])
    issues = check_workbook(wb, project)
    kinds = {i.kind for i in issues}
    assert 'sheet_file_missing_from_manifest' in kinds
