import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main
from sheetwright.source.reader import read_source
from tests.fixtures.external_xlsx import write_xlsx_with_external_ref


@contextmanager
def _empty_project():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        p = tmp_path / 'proj'
        p.mkdir()
        (p / 'sheetwright.toml').write_text(
            '[project]\nname = "ext"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "ext"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()
        yield tmp_path, p


@test
def import_errors_on_external_refs_without_flatten():
    with _empty_project() as (tmp_path, p):
        src = tmp_path / 'in.xlsx'
        write_xlsx_with_external_ref(src, cached_value=42.0)
        runner = CliRunner()
        r = runner.invoke(main, ['import', str(src), '--project', str(p)])
        assert r.exit_code != 0
        assert 'external reference' in r.output.lower()


@test
def import_flatten_replaces_formula_with_cached_value():
    with _empty_project() as (tmp_path, p):
        src = tmp_path / 'in.xlsx'
        write_xlsx_with_external_ref(src, cached_value=42.0)
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
