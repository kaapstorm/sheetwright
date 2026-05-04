import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from sheetwright.xlsx.reader import read_xlsx
from sheetwright.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_validation_xlsx


@contextmanager
def _round_tripped():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'v.xlsx'
        out = tmp_path / 'v-out.xlsx'
        write_validation_xlsx(src)
        wb = read_xlsx(src)
        write_xlsx(wb, out)
        yield read_xlsx(out)


@test
def list_validation_round_trips():
    with _round_tripped() as wb:
        sheet = wb.sheet('V')
        matches = [v for v in sheet.validations if v.type == 'list']
        assert matches, 'no list validation found'
        v = matches[0]
        assert 'A1:A10' in v.ranges
        assert v.formula1 == '"yes,no,maybe"'


@test
def whole_range_validation_round_trips():
    with _round_tripped() as wb:
        sheet = wb.sheet('V')
        matches = [v for v in sheet.validations if v.type == 'whole']
        assert matches, 'no whole-number validation found'
        v = matches[0]
        assert v.operator == 'between'
        assert v.formula1 == 1 or v.formula1 == '1'
        assert v.formula2 == 100 or v.formula2 == '100'
        assert 'B1:B10' in v.ranges
