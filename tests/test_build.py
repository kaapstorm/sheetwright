from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.xlsx.reader import read_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def imported_project(tmp_path):
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
    result = runner.invoke(main, ['import', str(src), '--project', str(p)])
    assert result.exit_code == 0, result.output
    yield p


@use(imported_project)
def test_build_produces_xlsx():
    p = imported_project()
    runner = CliRunner()
    result = runner.invoke(main, ['build', '--project', str(p)])
    assert result.exit_code == 0, result.output
    out_xlsx = p / 'build' / 'in.xlsx'
    assert out_xlsx.is_file()


@use(imported_project)
def test_built_xlsx_has_same_sheets_and_values():
    p = imported_project()
    runner = CliRunner()
    runner.invoke(main, ['build', '--project', str(p)])
    wb = read_xlsx(p / 'build' / 'in.xlsx')
    assert [s.name for s in wb.sheets] == ['Inputs', 'Outputs']
    assert wb.sheet('Inputs').get('B1').value == 0.04
    assert (
        wb.sheet('Outputs').get('B1').formula == '=Inputs!B2 * (1 + Inputs!B1)'
    )
