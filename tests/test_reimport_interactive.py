from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def populated_project(tmp_path: Path):
    """A project with source already populated by an earlier import."""
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
    runner = CliRunner()
    runner.invoke(main, ['import', str(src), '--project', str(p)])
    yield p, src


@use(populated_project)
def test_reimport_with_no_changes_reports_clean(tmp_path: Path):
    p, src = populated_project()
    runner = CliRunner()
    r = runner.invoke(main, ['import', str(src), '--project', str(p)])
    assert r.exit_code == 0, r.output
    assert 'no changes' in r.output.lower()


@use(populated_project)
def test_reimport_reject_leaves_source_unchanged(tmp_path: Path):
    p, src = populated_project()
    md = p / 'sheets' / '01_inputs.md'
    before = md.read_text()

    new_src = tmp_path / 'new.xlsx'
    import openpyxl

    wb = openpyxl.Workbook()
    s1 = wb.active
    s1.title = 'Inputs'
    s1['A1'] = 'growth_rate'
    s1['B1'] = 0.99
    s1['A2'] = 'base_revenue'
    s1['B2'] = 1_000_000
    wb.create_sheet('Outputs')
    wb.save(new_src)

    runner = CliRunner()
    r = runner.invoke(
        main,
        ['import', str(new_src), '--project', str(p)],
        input='r\n',
    )
    assert r.exit_code == 0, r.output
    assert md.read_text() == before


@use(populated_project)
def test_reimport_overwrite_replaces_source(tmp_path: Path):
    p, src = populated_project()
    md = p / 'sheets' / '01_inputs.md'

    new_src = tmp_path / 'new.xlsx'
    import openpyxl

    wb = openpyxl.Workbook()
    s1 = wb.active
    s1.title = 'Inputs'
    s1['A1'] = 'growth_rate'
    s1['B1'] = 0.99
    s1['A2'] = 'base_revenue'
    s1['B2'] = 1_000_000
    wb.create_sheet('Outputs')
    wb.save(new_src)

    runner = CliRunner()
    r = runner.invoke(
        main,
        ['import', str(new_src), '--project', str(p)],
        input='o\n',
    )
    assert r.exit_code == 0, r.output
    assert '0.99' in md.read_text()
