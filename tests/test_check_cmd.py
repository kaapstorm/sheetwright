from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main


@fixture
def project(tmp_path: Path):
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "x"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    yield p


@use(project)
def test_check_clean_returns_zero():
    p = project()
    runner = CliRunner()
    r = runner.invoke(main, ['check', '--project', str(p)])
    assert r.exit_code == 0
    assert 'no issues' in r.output.lower()


@use(project)
def test_check_finds_orphaned_sheet():
    p = project()
    (p / 'sheets' / '01_orphan.md').write_text(
        '| (cell) | A |\n| --- | --- |\n'
    )
    runner = CliRunner()
    r = runner.invoke(main, ['check', '--project', str(p)])
    assert r.exit_code != 0
    assert 'sheet_file_missing_from_manifest' in r.output
