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
    # workbook.toml stores DISPLAY NAMES; the source reader/writer
    # reconstruct stems via slugify+index. Mirror that here.
    project = _project(tmp_path, ['Inputs'])
    (project.sheets_dir / '01_inputs.md').write_text(
        '| (cell) | A |\n| --- | --- |\n'
    )
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    issues = check_workbook(wb, project)
    assert issues == []


def test_clean_after_real_import_has_no_issues(tmp_path: Path):
    """Regression: after `claudesheets import`, check_workbook
    must not false-positive on the manifest/stem mismatch.

    `claudesheets import` writes display names to workbook.toml and
    files at `<NN>_<slug(name)>.md`. check_workbook should reconcile.
    """
    from click.testing import CliRunner

    from claudesheets.cli import main
    from tests.fixtures.workbooks import write_simple_xlsx

    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "in"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    CliRunner().invoke(main, ['import', str(src), '--project', str(p)])

    project = Project.open(p)
    from claudesheets.source.reader import read_source

    wb = read_source(project.root)
    issues = check_workbook(wb, project)
    kinds = {i.kind for i in issues}
    assert 'manifest_sheet_missing' not in kinds
    assert 'sheet_file_missing_from_manifest' not in kinds


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
    # Manifest lists 'Missing' (display name); expected stem is
    # '01_missing'. No file exists → manifest_sheet_missing fires.
    project = _project(tmp_path, ['Missing'])
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
