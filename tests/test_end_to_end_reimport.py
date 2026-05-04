import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _baseline():
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
        runner = CliRunner()
        runner.invoke(main, ['import', str(src), '--project', str(p)])
        runner.invoke(main, ['build', '--project', str(p)])
        yield p, src


@test
@requires_libreoffice
def full_escape_hatch_loop():
    with _baseline() as (p, _orig_src):
        runner = CliRunner()

        edited = p / 'build' / 'in.xlsx'
        import openpyxl

        wb = openpyxl.load_workbook(edited)
        wb['Inputs']['B1'] = 0.10
        wb.save(edited)

        r = runner.invoke(main, ['snapshot', '--project', str(p)])
        assert 'modified externally' in r.output.lower()

        r = runner.invoke(
            main,
            ['import', str(edited), '-I', '--project', str(p)],
        )
        assert r.exit_code == 0
        assert 'Inputs!B1' in r.output
        assert (p / '.sheetwright' / 'reimport.json').is_file()

        r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
        assert r.exit_code == 0
        md = p / 'sheets' / '01_inputs.md'
        assert '0.1' in md.read_text()
