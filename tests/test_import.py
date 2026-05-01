from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.workbooks import write_formatted_xlsx, write_simple_xlsx


@fixture
def project(tmp_path):
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "x"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    yield p


@use(project)
def test_import_writes_source_files(tmp_path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    runner = CliRunner()
    result = runner.invoke(
        main, ['import', str(src), '--project', str(project())]
    )
    assert result.exit_code == 0, result.output
    assert (project() / 'sheets' / '01_inputs.md').is_file()
    assert (project() / 'sheets' / '02_outputs.md').is_file()
    text = (project() / 'workbook.toml').read_text()
    assert 'growth_rate' in text


@use(project)
def test_import_with_archive_copies_xlsx(tmp_path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    runner = CliRunner()
    result = runner.invoke(
        main,
        ['import', str(src), '--archive', '--project', str(project())],
    )
    assert result.exit_code == 0, result.output
    archives = list((project() / 'imports').glob('*.xlsx'))
    assert len(archives) == 1


@use(project)
def test_import_refuses_when_source_already_populated(tmp_path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    (project() / 'sheets' / '01_existing.md').write_text(
        '| (cell) | A |\n| --- | --- |\n'
    )

    runner = CliRunner()
    result = runner.invoke(
        main, ['import', str(src), '--project', str(project())]
    )
    assert result.exit_code != 0
    assert (
        'non-empty' in result.output.lower()
        or 'exist' in result.output.lower()
    )


@use(project)
def test_import_round_trips_formatted_workbook(tmp_path):
    src = tmp_path / 'fmt.xlsx'
    write_formatted_xlsx(src)
    runner = CliRunner()
    result = runner.invoke(
        main, ['import', str(src), '--project', str(project())]
    )
    assert result.exit_code == 0, result.output
    yamls = list((project() / 'sheets').glob('*.yaml'))
    assert yamls, 'expected a YAML sidecar for the formatted sheet'
