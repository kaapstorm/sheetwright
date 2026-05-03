from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def built(tmp_path: Path):
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


@use(built)
def test_no_warning_when_build_unchanged():
    p = built()
    runner = CliRunner()
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert 'modified externally' not in r.output


@use(built, requires_libreoffice)
def test_warn_when_build_changed_externally():
    p = built()
    built_xlsx = p / 'build' / 'in.xlsx'
    data = built_xlsx.read_bytes()
    built_xlsx.write_bytes(data + b' ')
    runner = CliRunner()
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert 'modified externally' in r.output.lower()


@use(built)
def test_no_warning_when_no_record_yet(tmp_path: Path):
    p = tmp_path / 'fresh'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "in"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    runner = CliRunner()
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert 'modified externally' not in r.output


@use(built)
def test_warn_when_workbook_renamed_orphans_hash():
    """Renaming claudesheets.toml.name should emit an orphan note."""
    p = built()
    cs_toml = p / 'claudesheets.toml'
    cs_toml.write_text(
        cs_toml.read_text().replace('name = "in"', 'name = "renamed"')
    )
    runner = CliRunner()
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert 'orphaned' in r.output.lower() or 'orphan' in r.output.lower()
