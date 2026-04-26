from unmagic import fixture, use

from claudesheets.xlsx.reader import read_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def simple_xlsx(tmp_path):
    path = tmp_path / 'simple.xlsx'
    write_simple_xlsx(path)
    yield path


@use(simple_xlsx)
def test_read_returns_workbook_with_two_sheets():
    wb = read_xlsx(simple_xlsx())
    assert wb.name == 'simple'
    assert [s.name for s in wb.sheets] == ['Inputs', 'Outputs']


@use(simple_xlsx)
def test_read_preserves_values():
    wb = read_xlsx(simple_xlsx())
    inputs = wb.sheet('Inputs')
    assert inputs.get('A1').value == 'growth_rate'
    assert inputs.get('B1').value == 0.04
    assert inputs.get('B2').value == 1_000_000


@use(simple_xlsx)
def test_read_preserves_formula():
    wb = read_xlsx(simple_xlsx())
    outputs = wb.sheet('Outputs')
    cell = outputs.get('B1')
    assert cell.value is None
    assert cell.formula == '=Inputs!B2 * (1 + Inputs!B1)'


@use(simple_xlsx)
def test_read_preserves_named_range():
    wb = read_xlsx(simple_xlsx())
    names = {nr.name: nr for nr in wb.named_ranges}
    assert 'growth_rate' in names
    assert names['growth_rate'].ref == 'Inputs!$B$1'
    assert names['growth_rate'].scope == 'workbook'
