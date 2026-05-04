"""Tests for calc-driven MCP tools: do_recalc, do_snapshot, do_test."""

import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from claudesheets.mcp.server import do_recalc, do_snapshot, do_test
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _built():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        p = tmp_path / 'proj'
        p.mkdir()
        (p / 'claudesheets.toml').write_text(
            '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "in"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()
        (p / 'tests').mkdir()
        from click.testing import CliRunner

        from claudesheets.cli import main

        CliRunner().invoke(main, ['import', str(src), '--project', str(p)])
        CliRunner().invoke(main, ['build', '--project', str(p)])
        yield p


@test
@requires_libreoffice
def recalc_tool_runs_engine_and_caches():
    with _built() as p:
        out = do_recalc(project=str(p), force=False)
        assert out['ok'] is True
        cache_files = list((p / '.claudesheets' / 'calc').glob('*.json'))
        assert len(cache_files) == 1


@test
@requires_libreoffice
def snapshot_tool_initializes_then_clean():
    with _built() as p:
        out1 = do_snapshot(project=str(p), update=False)
        assert out1['ok'] is True
        assert 'initialized' in out1['message'].lower()

        out2 = do_snapshot(project=str(p), update=False)
        assert out2['ok'] is True
        assert 'no changes' in out2['message'].lower()


@test
def do_test_tool_passes_when_user_tests_pass():
    with _built() as p:
        (p / 'tests' / '__init__.py').write_text('')
        (p / 'tests' / 'test_simple.py').write_text(
            'from testsweet import test\n\n'
            '@test\n'
            'def passes():\n'
            '    assert 1 + 1 == 2\n'
        )
        out = do_test(project=str(p), targets=[])
        assert out['passed'] is True


@test
def do_test_tool_reports_failure():
    with _built() as p:
        (p / 'tests' / '__init__.py').write_text('')
        (p / 'tests' / 'test_fail.py').write_text(
            'from testsweet import test\n\n'
            '@test\n'
            'def fails():\n'
            '    assert False\n'
        )
        out = do_test(project=str(p), targets=[])
        assert out['passed'] is False
