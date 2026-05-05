import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _staged_project():
    """A project where -I has already been run; session.json exists."""
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

        new_src = tmp_path / 'new.xlsx'
        import openpyxl

        wb = openpyxl.Workbook()
        s1 = wb.active
        s1.title = 'Inputs'
        s1['A1'] = 'growth_rate'
        s1['B1'] = 0.99
        s1['A2'] = 'base_revenue'
        s1['B2'] = 1_000_000
        wb.create_sheet('Outputs')
        wb.save(new_src)

        r = runner.invoke(
            main, ['import', str(new_src), '-I', '--project', str(p)]
        )
        assert r.exit_code == 0
        assert (p / '.sheetwright' / 'reimport.json').is_file()
        yield p


@test
def apply_writes_staged_changes_to_source():
    with _staged_project() as p:
        md = p / 'sheets' / '01_inputs.md'
        runner = CliRunner()
        r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
        assert r.exit_code == 0, r.output
        assert '0.99' in md.read_text()
        assert not (p / '.sheetwright' / 'reimport.json').is_file()


@test
def abort_clears_session_and_does_not_change_source():
    with _staged_project() as p:
        md = p / 'sheets' / '01_inputs.md'
        before = md.read_text()
        runner = CliRunner()
        r = runner.invoke(main, ['import', '--abort', '--project', str(p)])
        assert r.exit_code == 0
        assert md.read_text() == before
        assert not (p / '.sheetwright' / 'reimport.json').is_file()


@test
def apply_refuses_when_staged_xlsx_modified():
    with _staged_project() as p:
        import json

        sess = json.loads((p / '.sheetwright' / 'reimport.json').read_text())
        # staged_filename is the copy under .sheetwright/staged/ — corrupt it
        staged_copy = p / '.sheetwright' / 'staged' / sess['staged_filename']
        staged_copy.write_bytes(staged_copy.read_bytes() + b' ')

        runner = CliRunner()
        r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
        assert r.exit_code != 0
        assert 'modified' in r.output.lower() or 're-stage' in r.output.lower()


@test
def apply_without_session_errors():
    with _staged_project() as p:
        runner = CliRunner()
        runner.invoke(main, ['import', '--abort', '--project', str(p)])
        r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
        assert r.exit_code != 0
        assert 'no staged' in r.output.lower() or 'session' in r.output.lower()
