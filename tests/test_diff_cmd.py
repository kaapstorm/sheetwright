from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def imported_project(tmp_path: Path):
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
    runner.invoke(main, ['build', '--project', str(p)])
    yield p


@use(imported_project)
def test_diff_clean_when_source_matches_build():
    p = imported_project()
    runner = CliRunner()
    r = runner.invoke(main, ['diff', '--project', str(p)])
    assert r.exit_code == 0, r.output
    assert 'no changes' in r.output.lower()


@use(imported_project)
def test_diff_against_external_xlsx(tmp_path: Path):
    p = imported_project()
    other = tmp_path / 'other.xlsx'
    # write_simple_xlsx is the same content as the imported source, so
    # diff should be empty even though the build hash differs.
    write_simple_xlsx(other)
    runner = CliRunner()
    r = runner.invoke(
        main,
        ['diff', '--project', str(p), '--vs', f'xlsx:{other}'],
    )
    assert r.exit_code == 0, r.output


@use(imported_project)
def test_diff_reports_changed_cell():
    p = imported_project()
    # Mutate the source's Inputs!B1 in the markdown.
    md = p / 'sheets' / '01_inputs.md'
    text = md.read_text()
    md.write_text(text.replace('0.04', '0.10'))
    runner = CliRunner()
    r = runner.invoke(main, ['diff', '--project', str(p)])
    # Output is non-empty diff; exit code is non-zero (diffs found).
    assert r.exit_code != 0
    assert 'Inputs!B1' in r.output
    assert '0.04' in r.output
    assert '0.1' in r.output  # rendered as 0.1 by repr()
