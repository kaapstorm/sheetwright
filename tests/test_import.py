import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main
from tests.fixtures.workbooks import write_formatted_xlsx, write_simple_xlsx


@contextmanager
def _project():
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
        yield tmp_path, p


@test
def import_writes_source_files():
    with _project() as (tmp_path, p):
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        runner = CliRunner()
        result = runner.invoke(main, ['import', str(src), '--project', str(p)])
        assert result.exit_code == 0, result.output
        assert (p / 'sheets' / '01_inputs.md').is_file()
        assert (p / 'sheets' / '02_outputs.md').is_file()
        text = (p / 'workbook.toml').read_text()
        assert 'growth_rate' in text


@test
def import_with_archive_copies_xlsx():
    with _project() as (tmp_path, p):
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        runner = CliRunner()
        result = runner.invoke(
            main,
            ['import', str(src), '--archive', '--project', str(p)],
        )
        assert result.exit_code == 0, result.output
        archives = list((p / 'imports').glob('*.xlsx'))
        assert len(archives) == 1


@test
def import_enters_reimport_flow_when_source_populated():
    with _project() as (tmp_path, p):
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        (p / 'sheets' / '01_existing.md').write_text(
            '| (cell) | A |\n| --- | --- |\n'
        )

        runner = CliRunner()
        result = runner.invoke(
            main,
            ['import', str(src), '--project', str(p)],
            input='r\n',
        )
        assert result.exit_code == 0, result.output
        assert 'rejected' in result.output.lower()


@test
def import_round_trips_formatted_workbook():
    with _project() as (tmp_path, p):
        src = tmp_path / 'fmt.xlsx'
        write_formatted_xlsx(src)
        runner = CliRunner()
        result = runner.invoke(main, ['import', str(src), '--project', str(p)])
        assert result.exit_code == 0, result.output
        yamls = list((p / 'sheets').glob('*.yaml'))
        assert yamls, 'expected a YAML sidecar for the formatted sheet'
