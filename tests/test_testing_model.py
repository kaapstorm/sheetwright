from pathlib import Path

import pytest
from unmagic import fixture, use

from claudesheets.calc.base import CalcEngine, CalcResult
from claudesheets.model.cell import Cell
from claudesheets.model.workbook import Sheet, Workbook
from claudesheets.testing import Model
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def project(tmp_path: Path):
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
    from click.testing import CliRunner
    from claudesheets.cli import main

    r = CliRunner().invoke(main, ['import', str(src), '--project', str(p)])
    assert r.exit_code == 0
    yield p


@use(project, requires_libreoffice)
def test_model_get_returns_calculated_value():
    m = Model.open(project())
    assert m.get('Outputs!B1') == 1_040_000


@use(project, requires_libreoffice)
def test_model_set_invalidates_cache_and_recalculates():
    m = Model.open(project())
    m.set('Inputs!B1', 0.10)
    assert m.get('Outputs!B1') == 1_100_000


@use(project, requires_libreoffice)
def test_model_set_resolves_named_range():
    m = Model.open(project())
    m.set('growth_rate', 0.20)
    assert m.get('Outputs!B1') == 1_200_000


@use(project, requires_libreoffice)
def test_model_get_literal_value():
    m = Model.open(project())
    assert m.get('Inputs!B2') == 1_000_000


class _EmptyEngine(CalcEngine):
    def evaluate(self, xlsx_path: Path) -> CalcResult:
        return {}


def test_model_get_raises_when_formula_cell_missing_from_result(
    tmp_path: Path,
):
    wb = Workbook(name='x', sheets=[Sheet(name='S')])
    wb.sheet('S').set('A1', Cell(formula='=1+1'))
    m = Model(wb, _EmptyEngine())
    with pytest.raises(RuntimeError, match='S!A1'):
        m.get('S!A1')


def test_model_get_returns_literal_when_no_formula(tmp_path: Path):
    wb = Workbook(name='x', sheets=[Sheet(name='S')])
    wb.sheet('S').set('A1', Cell(value='hello'))
    m = Model(wb, _EmptyEngine())
    assert m.get('S!A1') == 'hello'
