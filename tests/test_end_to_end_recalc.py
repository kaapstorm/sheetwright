from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
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
    (p / 'tests').mkdir()
    runner = CliRunner()
    runner.invoke(main, ['import', str(src), '--project', str(p)])
    runner.invoke(main, ['build', '--project', str(p)])
    yield p


@use(project, requires_libreoffice)
def test_full_recalc_then_snapshot_then_model_set():
    p = project()
    runner = CliRunner()

    r = runner.invoke(main, ['recalc', '--project', str(p)])
    assert r.exit_code == 0, r.output

    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code == 0, r.output

    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code == 0, r.output
    assert 'no changes' in r.output.lower()

    m = Model.open(p)
    assert m.get('Outputs!B1') == 1_040_000
    m.set('Inputs!B1', 0.10)
    assert m.get('Outputs!B1') == 1_100_000
