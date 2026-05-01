from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main


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


@fixture
def project_with_tests(tmp_path: Path):
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "x"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    (p / 'tests').mkdir()
    (p / 'tests' / '__init__.py').write_text('')
    (p / 'tests' / 'test_simple.py').write_text(_PASS_TEST)
    yield p


@use(project_with_tests)
def test_test_command_runs_testsweet():
    p = project_with_tests()
    runner = CliRunner()
    r = runner.invoke(main, ['test', '--project', str(p)])
    assert r.exit_code == 0, r.output
    assert 'ok' in r.output
    assert 'passes' in r.output


@use(project_with_tests)
def test_test_command_propagates_failure_exit_code():
    p = project_with_tests()
    (p / 'tests' / 'test_fail.py').write_text(_FAIL_TEST)
    runner = CliRunner()
    r = runner.invoke(main, ['test', '--project', str(p)])
    assert r.exit_code != 0
    assert 'FAIL' in r.output


@use(project_with_tests)
def test_test_command_supports_target_selection():
    p = project_with_tests()
    (p / 'tests' / 'test_other.py').write_text(_FAIL_TEST)
    runner = CliRunner()
    # Pass a specific target so we only run the passing test.
    r = runner.invoke(
        main,
        ['test', '--project', str(p), 'tests/test_simple.py'],
    )
    assert r.exit_code == 0, r.output


@use(project_with_tests)
def test_test_command_errors_when_tests_dir_missing(tmp_path: Path):
    p = tmp_path / 'no-tests'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "x"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    runner = CliRunner()
    r = runner.invoke(main, ['test', '--project', str(p)])
    assert r.exit_code != 0
    assert 'tests/' in r.output or 'tests' in r.output.lower()
