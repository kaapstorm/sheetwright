from unmagic import fixture, use

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_validation_xlsx


@fixture
def round_tripped(tmp_path):
    src = tmp_path / 'v.xlsx'
    out = tmp_path / 'v-out.xlsx'
    write_validation_xlsx(src)
    wb = read_xlsx(src)
    write_xlsx(wb, out)
    yield read_xlsx(out)


@use(round_tripped)
def test_list_validation_round_trips():
    sheet = round_tripped().sheet('V')
    matches = [v for v in sheet.validations if v.type == 'list']
    assert matches, 'no list validation found'
    v = matches[0]
    assert 'A1:A10' in v.ranges
    assert v.formula1 == '"yes,no,maybe"'


@use(round_tripped)
def test_whole_range_validation_round_trips():
    sheet = round_tripped().sheet('V')
    matches = [v for v in sheet.validations if v.type == 'whole']
    assert matches, 'no whole-number validation found'
    v = matches[0]
    assert v.operator == 'between'
    assert v.formula1 == 1 or v.formula1 == '1'
    assert v.formula2 == 100 or v.formula2 == '100'
    assert 'B1:B10' in v.ranges
