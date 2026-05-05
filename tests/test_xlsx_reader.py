import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from sheetwright.xlsx.reader import read_xlsx
from tests.fixtures.workbooks import write_simple_xlsx
from sheetwright.security import SecurityLimits


@contextmanager
def _simple_xlsx():
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / 'simple.xlsx'
        write_simple_xlsx(path)
        yield path


@test
def read_returns_workbook_with_two_sheets():
    with _simple_xlsx() as path:
        wb = read_xlsx(path, limits=SecurityLimits.defaults())
        assert wb.name == 'simple'
        assert [s.name for s in wb.sheets] == ['Inputs', 'Outputs']


@test
def read_preserves_values():
    with _simple_xlsx() as path:
        wb = read_xlsx(path, limits=SecurityLimits.defaults())
        inputs = wb.sheet('Inputs')
        assert inputs.get('A1').value == 'growth_rate'
        assert inputs.get('B1').value == 0.04
        assert inputs.get('B2').value == 1_000_000


@test
def read_preserves_formula():
    with _simple_xlsx() as path:
        wb = read_xlsx(path, limits=SecurityLimits.defaults())
        outputs = wb.sheet('Outputs')
        cell = outputs.get('B1')
        assert cell.value is None
        assert cell.formula == '=Inputs!B2 * (1 + Inputs!B1)'


@test
def read_preserves_named_range():
    with _simple_xlsx() as path:
        wb = read_xlsx(path, limits=SecurityLimits.defaults())
        names = {nr.name: nr for nr in wb.named_ranges}
        assert 'growth_rate' in names
        assert names['growth_rate'].ref == 'Inputs!$B$1'
        assert names['growth_rate'].scope == 'workbook'
