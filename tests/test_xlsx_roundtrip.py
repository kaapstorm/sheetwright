from unmagic import fixture, use

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def round_tripped(tmp_path):
    src = tmp_path / 'in.xlsx'
    out = tmp_path / 'out.xlsx'
    write_simple_xlsx(src)
    wb = read_xlsx(src)
    write_xlsx(wb, out)
    yield read_xlsx(out)


@use(round_tripped)
def test_round_trip_preserves_sheets():
    wb = round_tripped()
    assert [s.name for s in wb.sheets] == ['Inputs', 'Outputs']


@use(round_tripped)
def test_round_trip_preserves_values():
    wb = round_tripped()
    assert wb.sheet('Inputs').get('B1').value == 0.04
    assert wb.sheet('Inputs').get('B2').value == 1_000_000


@use(round_tripped)
def test_round_trip_preserves_formula():
    wb = round_tripped()
    cell = wb.sheet('Outputs').get('B1')
    assert cell.formula == '=Inputs!B2 * (1 + Inputs!B1)'


@use(round_tripped)
def test_round_trip_preserves_named_range():
    wb = round_tripped()
    names = {nr.name: nr for nr in wb.named_ranges}
    assert names['growth_rate'].ref == 'Inputs!$B$1'
