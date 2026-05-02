import subprocess
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.workbooks import write_simple_xlsx


def _git(*args: str, cwd: Path) -> None:
    subprocess.run(
        ['git', *args],
        cwd=cwd,
        check=True,
        capture_output=True,
    )


@fixture
def gitted_project(tmp_path: Path):
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
    _git('init', cwd=p)
    _git('config', 'user.email', 'test@example.com', cwd=p)
    _git('config', 'user.name', 'Tester', cwd=p)
    runner = CliRunner()
    runner.invoke(main, ['import', str(src), '--project', str(p)])
    _git('add', '-A', cwd=p)
    _git('commit', '-m', 'initial', cwd=p)
    yield p, src


@use(gitted_project)
def test_reimport_blocked_when_source_has_uncommitted_changes(tmp_path: Path):
    p, src = gitted_project()
    # Mutate a sheet markdown so it's dirty.
    md = p / 'sheets' / '01_inputs.md'
    md.write_text(md.read_text() + '\n')

    runner = CliRunner()
    r = runner.invoke(main, ['import', str(src), '--project', str(p)])
    assert r.exit_code != 0
    assert 'uncommitted' in r.output.lower()


@use(gitted_project)
def test_reimport_allowed_with_force(tmp_path: Path):
    p, src = gitted_project()
    md = p / 'sheets' / '01_inputs.md'
    md.write_text(md.read_text() + '\n')

    runner = CliRunner()
    r = runner.invoke(
        main,
        ['import', str(src), '--force', '--project', str(p)],
        input='r\n',  # answer the prompt with reject so we don't actually overwrite
    )
    # --force should suppress the uncommitted-changes block; the diff
    # then runs and we reject. exit code 0.
    assert r.exit_code == 0


def test_no_git_no_guard(tmp_path: Path):
    """If the project isn't a git repo, the guard is a no-op."""
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

    # Re-import without git; should not trip the guard.
    md = p / 'sheets' / '01_inputs.md'
    md.write_text(md.read_text() + '\n')
    r = runner.invoke(
        main,
        ['import', str(src), '--project', str(p)],
        input='r\n',
    )
    assert r.exit_code == 0
