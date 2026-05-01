from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def imported(tmp_path: Path):
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
    r = runner.invoke(main, ['import', str(src), '--project', str(p)])
    assert r.exit_code == 0, r.output
    r = runner.invoke(main, ['build', '--project', str(p)])
    assert r.exit_code == 0, r.output
    yield p


@use(imported, requires_libreoffice)
def test_recalc_writes_cache_entry():
    p = imported()
    runner = CliRunner()
    r = runner.invoke(main, ['recalc', '--project', str(p)])
    assert r.exit_code == 0, r.output
    cache_files = list((p / '.claudesheets' / 'calc').glob('*.json'))
    assert len(cache_files) == 1


@use(imported, requires_libreoffice)
def test_recalc_is_idempotent_uses_cache():
    from claudesheets.commands.recalc_cmd import CACHE_HIT_MESSAGE

    p = imported()
    runner = CliRunner()
    runner.invoke(main, ['recalc', '--project', str(p)])
    out = runner.invoke(main, ['recalc', '--project', str(p)])
    assert out.exit_code == 0
    assert CACHE_HIT_MESSAGE in out.output


@use(imported, requires_libreoffice)
def test_recalc_force_rebuilds_cache():
    from claudesheets.commands.recalc_cmd import CACHE_HIT_MESSAGE

    p = imported()
    runner = CliRunner()
    runner.invoke(main, ['recalc', '--project', str(p)])
    out = runner.invoke(main, ['recalc', '--project', str(p), '--force'])
    assert out.exit_code == 0
    assert CACHE_HIT_MESSAGE not in out.output
