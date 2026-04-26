from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main


@fixture
def tmp_project(tmp_path):
    yield tmp_path / 'my-model'


@use(tmp_project)
def test_init_creates_project_skeleton():
    project = tmp_project()
    runner = CliRunner()
    result = runner.invoke(main, ['init', str(project)])
    assert result.exit_code == 0, result.output

    assert (project / 'claudesheets.toml').is_file()
    assert (project / 'workbook.toml').is_file()
    assert (project / 'sheets').is_dir()
    assert (project / 'data').is_dir()
    assert (project / '.gitignore').is_file()

    gitignore = (project / '.gitignore').read_text()
    assert 'build/' in gitignore
    assert '.claudesheets/' in gitignore


@use(tmp_project)
def test_init_refuses_non_empty_directory():
    project = tmp_project()
    project.mkdir()
    (project / 'stuff.txt').write_text('hi')

    runner = CliRunner()
    result = runner.invoke(main, ['init', str(project)])
    assert result.exit_code != 0
    assert 'not empty' in result.output.lower()
