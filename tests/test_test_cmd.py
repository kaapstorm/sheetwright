import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main


_PASS_TEST = """\
from testsweet import test


@test
def passes():
    assert 1 + 1 == 2
"""

_FAIL_TEST = """\
from testsweet import test


@test
def fails():
    assert False
"""


@contextmanager
def _project_with_tests():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        p = tmp_path / 'proj'
        p.mkdir()
        (p / 'sheetwright.toml').write_text(
            '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "x"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()
        (p / 'tests').mkdir()
        (p / 'tests' / '__init__.py').write_text('')
        (p / 'tests' / 'test_simple.py').write_text(_PASS_TEST)
        yield tmp_path, p


@test
def test_command_runs_testsweet():
    with _project_with_tests() as (_tmp, p):
        runner = CliRunner()
        r = runner.invoke(main, ['test', '--project', str(p)])
        assert r.exit_code == 0, r.output
        assert 'ok' in r.output
        assert 'passes' in r.output


@test
def test_command_propagates_failure_exit_code():
    with _project_with_tests() as (_tmp, p):
        (p / 'tests' / 'test_fail.py').write_text(_FAIL_TEST)
        runner = CliRunner()
        r = runner.invoke(main, ['test', '--project', str(p)])
        assert r.exit_code != 0
        assert 'FAIL' in r.output


@test
def test_command_supports_target_selection():
    with _project_with_tests() as (_tmp, p):
        (p / 'tests' / 'test_other.py').write_text(_FAIL_TEST)
        runner = CliRunner()
        r = runner.invoke(
            main,
            ['test', '--project', str(p), 'tests/test_simple.py'],
        )
        assert r.exit_code == 0, r.output
        assert 'passes' in r.output
        assert 'ok' in r.output
        assert 'fails' not in r.output


@test
def test_command_rejects_path_outside_project():
    with _project_with_tests() as (_tmp, p):
        runner = CliRunner()
        r = runner.invoke(
            main,
            ['test', '--project', str(p), '../escape.py'],
        )
        assert r.exit_code != 0
        assert 'Traceback' not in r.output
        assert 'escape' in r.output or 'outside' in r.output


@test
def test_command_errors_when_tests_dir_missing():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        p = tmp_path / 'no-tests'
        p.mkdir()
        (p / 'sheetwright.toml').write_text(
            '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "x"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()
        runner = CliRunner()
        r = runner.invoke(main, ['test', '--project', str(p)])
        assert r.exit_code != 0
        assert 'tests/' in r.output or 'tests' in r.output.lower()
