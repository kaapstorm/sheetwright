import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _populated_project():
    """A project with source already populated by an earlier import."""
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
        yield tmp_path, p, src


def _write_modified_xlsx(path: Path, b1_value: float = 0.99) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    s1 = wb.active
    s1.title = 'Inputs'
    s1['A1'] = 'growth_rate'
    s1['B1'] = b1_value
    s1['A2'] = 'base_revenue'
    s1['B2'] = 1_000_000
    wb.create_sheet('Outputs')
    wb.save(path)


@test
def reimport_with_no_changes_reports_clean():
    with _populated_project() as (_tmp, p, src):
        runner = CliRunner()
        r = runner.invoke(main, ['import', str(src), '--project', str(p)])
        assert r.exit_code == 0, r.output
        assert 'no changes' in r.output.lower()


@test
def reimport_reject_leaves_source_unchanged():
    with _populated_project() as (tmp_path, p, _src):
        md = p / 'sheets' / '01_inputs.md'
        before = md.read_text()

        new_src = tmp_path / 'new.xlsx'
        _write_modified_xlsx(new_src)

        runner = CliRunner()
        r = runner.invoke(
            main,
            ['import', str(new_src), '--project', str(p)],
            input='r\n',
        )
        assert r.exit_code == 0, r.output
        assert md.read_text() == before


@test
def reimport_overwrite_replaces_source():
    with _populated_project() as (tmp_path, p, _src):
        md = p / 'sheets' / '01_inputs.md'

        new_src = tmp_path / 'new.xlsx'
        _write_modified_xlsx(new_src)

        runner = CliRunner()
        r = runner.invoke(
            main,
            ['import', str(new_src), '--project', str(p)],
            input='o\n',
        )
        assert r.exit_code == 0, r.output
        assert '0.99' in md.read_text()
