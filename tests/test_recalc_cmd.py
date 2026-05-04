import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _imported():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        p = tmp_path / 'proj'
        p.mkdir()
        (p / 'sheetwright.toml').write_text(
            '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "in"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()
        runner = CliRunner()
        r = runner.invoke(main, ['import', str(src), '--project', str(p)])
        assert r.exit_code == 0, r.output
        r = runner.invoke(main, ['build', '--project', str(p)])
        assert r.exit_code == 0, r.output
        yield p


@test
@requires_libreoffice
def recalc_writes_cache_entry():
    with _imported() as p:
        runner = CliRunner()
        r = runner.invoke(main, ['recalc', '--project', str(p)])
        assert r.exit_code == 0, r.output
        cache_files = list((p / '.sheetwright' / 'calc').glob('*.json'))
        assert len(cache_files) == 1


@test
@requires_libreoffice
def recalc_is_idempotent_uses_cache():
    from sheetwright.commands.recalc_cmd import CACHE_HIT_MESSAGE

    with _imported() as p:
        runner = CliRunner()
        runner.invoke(main, ['recalc', '--project', str(p)])
        out = runner.invoke(main, ['recalc', '--project', str(p)])
        assert out.exit_code == 0
        assert CACHE_HIT_MESSAGE in out.output


@test
@requires_libreoffice
def recalc_force_rebuilds_cache():
    from sheetwright.commands.recalc_cmd import CACHE_HIT_MESSAGE

    with _imported() as p:
        runner = CliRunner()
        runner.invoke(main, ['recalc', '--project', str(p)])
        out = runner.invoke(main, ['recalc', '--project', str(p), '--force'])
        assert out.exit_code == 0
        assert CACHE_HIT_MESSAGE not in out.output
