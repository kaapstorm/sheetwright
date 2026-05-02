from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.source.reader import read_source
from tests.fixtures.external_xlsx import write_xlsx_with_external_ref


@fixture
def empty_project(tmp_path: Path):
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "ext"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "ext"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    yield p


@use(empty_project)
def test_import_errors_on_external_refs_without_flatten(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_xlsx_with_external_ref(src, cached_value=42.0)
    p = empty_project()
    runner = CliRunner()
    r = runner.invoke(main, ['import', str(src), '--project', str(p)])
    assert r.exit_code != 0
    assert 'external reference' in r.output.lower()


@use(empty_project)
def test_import_flatten_replaces_formula_with_cached_value(
    tmp_path: Path,
):
    src = tmp_path / 'in.xlsx'
    write_xlsx_with_external_ref(src, cached_value=42.0)
    p = empty_project()
    runner = CliRunner()
    r = runner.invoke(
        main,
        ['import', str(src), '--flatten', '--project', str(p)],
    )
    assert r.exit_code == 0, r.output
    wb = read_source(p)
    cell = wb.sheet('S').get('A1')
    assert cell.formula is None
    assert cell.value == 42.0
