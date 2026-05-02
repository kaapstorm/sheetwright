from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def staged_project(tmp_path: Path):
    """A project where -I has already been run; session.json exists."""
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

    # Build a different xlsx and stage it via -I.
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

    r = runner.invoke(
        main, ['import', str(new_src), '-I', '--project', str(p)]
    )
    assert r.exit_code == 0
    assert (p / '.claudesheets' / 'reimport.json').is_file()
    yield p


@use(staged_project)
def test_apply_writes_staged_changes_to_source():
    p = staged_project()
    md = p / 'sheets' / '01_inputs.md'
    runner = CliRunner()
    r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
    assert r.exit_code == 0, r.output
    assert '0.99' in md.read_text()
    assert not (p / '.claudesheets' / 'reimport.json').is_file()


@use(staged_project)
def test_abort_clears_session_and_does_not_change_source():
    p = staged_project()
    md = p / 'sheets' / '01_inputs.md'
    before = md.read_text()
    runner = CliRunner()
    r = runner.invoke(main, ['import', '--abort', '--project', str(p)])
    assert r.exit_code == 0
    assert md.read_text() == before
    assert not (p / '.claudesheets' / 'reimport.json').is_file()


@use(staged_project)
def test_apply_without_session_errors():
    p = staged_project()
    runner = CliRunner()
    runner.invoke(main, ['import', '--abort', '--project', str(p)])
    r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
    assert r.exit_code != 0
    assert 'no staged' in r.output.lower() or 'session' in r.output.lower()
