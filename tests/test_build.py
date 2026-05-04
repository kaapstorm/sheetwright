import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main
from sheetwright.xlsx.reader import read_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _imported_project():
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
        result = runner.invoke(main, ['import', str(src), '--project', str(p)])
        assert result.exit_code == 0, result.output
        yield p


@test
def build_produces_xlsx():
    with _imported_project() as p:
        runner = CliRunner()
        result = runner.invoke(main, ['build', '--project', str(p)])
        assert result.exit_code == 0, result.output
        out_xlsx = p / 'build' / 'in.xlsx'
        assert out_xlsx.is_file()


@test
def built_xlsx_has_same_sheets_and_values():
    with _imported_project() as p:
        runner = CliRunner()
        runner.invoke(main, ['build', '--project', str(p)])
        wb = read_xlsx(p / 'build' / 'in.xlsx')
        assert [s.name for s in wb.sheets] == ['Inputs', 'Outputs']
        assert wb.sheet('Inputs').get('B1').value == 0.04
        assert (
            wb.sheet('Outputs').get('B1').formula
            == '=Inputs!B2 * (1 + Inputs!B1)'
        )
