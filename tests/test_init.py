import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from claudesheets.cli import main


@contextmanager
def _tmp_project():
    with tempfile.TemporaryDirectory() as td:
        yield Path(td) / 'my-model'


@test
def init_creates_project_skeleton():
    with _tmp_project() as project:
        runner = CliRunner()
        result = runner.invoke(main, ['init', str(project)])
        assert result.exit_code == 0, result.output

        assert (project / 'claudesheets.toml').is_file()
        assert (project / 'workbook.toml').is_file()
        assert (project / 'sheets').is_dir()
        assert (project / 'data').is_dir()
        assert (project / 'tests').is_dir()
        assert (project / 'tests' / '__init__.py').is_file()
        assert (project / '.gitignore').is_file()

        gitignore = (project / '.gitignore').read_text()
        assert 'build/' in gitignore
        assert '.claudesheets/' in gitignore


@test
def init_refuses_non_empty_directory():
    with _tmp_project() as project:
        project.mkdir()
        (project / 'stuff.txt').write_text('hi')

        runner = CliRunner()
        result = runner.invoke(main, ['init', str(project)])
        assert result.exit_code != 0
        assert 'not empty' in result.output.lower()
