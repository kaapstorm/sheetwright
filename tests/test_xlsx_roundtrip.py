import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _round_tripped():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        out = tmp_path / 'out.xlsx'
        write_simple_xlsx(src)
        wb = read_xlsx(src)
        write_xlsx(wb, out)
        yield read_xlsx(out)


@test
def round_trip_preserves_sheets():
    with _round_tripped() as wb:
        assert [s.name for s in wb.sheets] == ['Inputs', 'Outputs']


@test
def round_trip_preserves_values():
    with _round_tripped() as wb:
        assert wb.sheet('Inputs').get('B1').value == 0.04
        assert wb.sheet('Inputs').get('B2').value == 1_000_000


@test
def round_trip_preserves_formula():
    with _round_tripped() as wb:
        cell = wb.sheet('Outputs').get('B1')
        assert cell.formula == '=Inputs!B2 * (1 + Inputs!B1)'


@test
def round_trip_preserves_named_range():
    with _round_tripped() as wb:
        names = {nr.name: nr for nr in wb.named_ranges}
        assert names['growth_rate'].ref == 'Inputs!$B$1'
