import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import catch_exceptions, test

from sheetwright.calc.base import CalcEngine, CalcResult
from sheetwright.model.cell import Cell
from sheetwright.model.workbook import Sheet, Workbook
from sheetwright.testing import Model
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _project():
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
        from click.testing import CliRunner
        from sheetwright.cli import main

        r = CliRunner().invoke(main, ['import', str(src), '--project', str(p)])
        assert r.exit_code == 0
        yield p


@test
@requires_libreoffice
def model_get_returns_calculated_value():
    with _project() as p:
        m = Model.open(p)
        assert m.get('Outputs!B1') == 1_040_000


@test
@requires_libreoffice
def model_set_invalidates_cache_and_recalculates():
    with _project() as p:
        m = Model.open(p)
        m.set('Inputs!B1', 0.10)
        assert m.get('Outputs!B1') == 1_100_000


@test
@requires_libreoffice
def model_set_resolves_named_range():
    with _project() as p:
        m = Model.open(p)
        m.set('growth_rate', 0.20)
        assert m.get('Outputs!B1') == 1_200_000


@test
@requires_libreoffice
def model_get_literal_value():
    with _project() as p:
        m = Model.open(p)
        assert m.get('Inputs!B2') == 1_000_000


class _EmptyEngine(CalcEngine):
    def evaluate(self, xlsx_path: Path) -> CalcResult:
        return {}


@test
def model_get_raises_when_formula_cell_missing_from_result():
    wb = Workbook(name='x', sheets=[Sheet(name='S')])
    wb.sheet('S').set('A1', Cell(formula='=1+1'))
    m = Model(wb, _EmptyEngine())
    with catch_exceptions() as excs:
        m.get('S!A1')
    assert excs and isinstance(excs[0], RuntimeError)
    assert 'S!A1' in str(excs[0])


@test
def model_get_returns_literal_when_no_formula():
    wb = Workbook(name='x', sheets=[Sheet(name='S')])
    wb.sheet('S').set('A1', Cell(value='hello'))
    m = Model(wb, _EmptyEngine())
    assert m.get('S!A1') == 'hello'
