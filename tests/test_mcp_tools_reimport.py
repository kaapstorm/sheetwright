import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import catch_exceptions, test

from claudesheets.mcp.errors import MCPError
from claudesheets.mcp.server import (
    do_reimport_abort,
    do_reimport_apply,
    do_reimport_stage,
)
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _populated():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        p = tmp_path / 'proj'
        p.mkdir()
        (p / 'claudesheets.toml').write_text(
            '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "in"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()
        from click.testing import CliRunner

        from claudesheets.cli import main

        CliRunner().invoke(main, ['import', str(src), '--project', str(p)])
        yield tmp_path, p, src


def _write_modified(path: Path, b1: float = 0.99) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    s = wb.active
    s.title = 'Inputs'
    s['A1'], s['B1'] = 'growth_rate', b1
    s['A2'], s['B2'] = 'base_revenue', 1_000_000
    wb.create_sheet('Outputs')
    wb.save(path)


@test
def reimport_stage_returns_empty_for_unchanged_xlsx():
    with _populated() as (_tmp, p, src):
        out = do_reimport_stage(
            xlsx=str(src), project=str(p), flatten=False, force=False
        )
        assert out['is_empty'] is True
        assert not (p / '.claudesheets' / 'reimport.json').is_file()


@test
def reimport_stage_then_apply():
    with _populated() as (tmp_path, p, _src):
        new_src = tmp_path / 'changed.xlsx'
        _write_modified(new_src)

        staged = do_reimport_stage(
            xlsx=str(new_src), project=str(p), flatten=False, force=False
        )
        assert staged['is_empty'] is False
        assert (p / '.claudesheets' / 'reimport.json').is_file()

        out = do_reimport_apply(project=str(p), archive=False)
        assert out['ok'] is True
        md = p / 'sheets' / '01_inputs.md'
        assert '0.99' in md.read_text()
        assert not (p / '.claudesheets' / 'reimport.json').is_file()


@test
def reimport_abort_clears_session():
    with _populated() as (tmp_path, p, _src):
        new_src = tmp_path / 'changed.xlsx'
        _write_modified(new_src, b1=0.50)

        do_reimport_stage(
            xlsx=str(new_src), project=str(p), flatten=False, force=False
        )
        out = do_reimport_abort(project=str(p))
        assert out['ok'] is True
        assert not (p / '.claudesheets' / 'reimport.json').is_file()


@test
def reimport_stage_then_restage_with_no_changes_clears_session():
    """C2 regression: a no-changes restage must clear any prior session."""
    with _populated() as (tmp_path, p, src):
        new_src = tmp_path / 'changed.xlsx'
        import openpyxl

        wb = openpyxl.Workbook()
        s = wb.active
        s.title = 'Inputs'
        s['A1'], s['B1'] = 'growth_rate', 0.99
        s['A2'], s['B2'] = 'base_revenue', 1_000_000
        s2 = wb.create_sheet('Outputs')
        s2['A1'] = 'revenue_2027'
        s2['B1'] = '=Inputs!B2 * (1 + Inputs!B1)'
        wb.save(new_src)

        do_reimport_stage(
            xlsx=str(new_src), project=str(p), flatten=False, force=False
        )
        assert (p / '.claudesheets' / 'reimport.json').is_file()

        out = do_reimport_stage(
            xlsx=str(src), project=str(p), flatten=False, force=False
        )
        assert out['is_empty'] is True
        assert not (p / '.claudesheets' / 'reimport.json').is_file()


@test
def reimport_apply_without_session_raises_typed_error():
    with _populated() as (_tmp, p, _src):
        with catch_exceptions() as excs:
            do_reimport_apply(project=str(p), archive=False)
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code == 'no_staged_session'
