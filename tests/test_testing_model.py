from pathlib import Path

from unmagic import fixture, use

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
