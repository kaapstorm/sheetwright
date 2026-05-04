import subprocess
import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from claudesheets.cli import main
from tests.fixtures.workbooks import write_simple_xlsx


def _git(*args: str, cwd: Path) -> None:
    subprocess.run(
        ['git', *args],
        cwd=cwd,
        check=True,
        capture_output=True,
    )


@contextmanager
def _gitted_project():
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
        _git('init', cwd=p)
        _git('config', 'user.email', 'test@example.com', cwd=p)
        _git('config', 'user.name', 'Tester', cwd=p)
        runner = CliRunner()
        runner.invoke(main, ['import', str(src), '--project', str(p)])
        _git('add', '-A', cwd=p)
        _git('commit', '-m', 'initial', cwd=p)
        yield p, src


@test
def reimport_blocked_when_source_has_uncommitted_changes():
    with _gitted_project() as (p, src):
        md = p / 'sheets' / '01_inputs.md'
        md.write_text(md.read_text() + '\n')

        runner = CliRunner()
        r = runner.invoke(main, ['import', str(src), '--project', str(p)])
        assert r.exit_code != 0
        assert 'uncommitted' in r.output.lower()


@test
def reimport_allowed_with_force():
    with _gitted_project() as (p, src):
        md = p / 'sheets' / '01_inputs.md'
        md.write_text(md.read_text() + '\n')

        runner = CliRunner()
        r = runner.invoke(
            main,
            ['import', str(src), '--force', '--project', str(p)],
            input='r\n',
        )
        assert r.exit_code == 0


@test
def no_git_no_guard():
    """If the project isn't a git repo, the guard is a no-op."""
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
        runner = CliRunner()
        runner.invoke(main, ['import', str(src), '--project', str(p)])

        md = p / 'sheets' / '01_inputs.md'
        md.write_text(md.read_text() + '\n')
        r = runner.invoke(
            main,
            ['import', str(src), '--project', str(p)],
            input='r\n',
        )
        assert r.exit_code == 0
