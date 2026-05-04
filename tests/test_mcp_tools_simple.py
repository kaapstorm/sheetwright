import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import catch_exceptions, test

from claudesheets.mcp.errors import MCPError
from claudesheets.mcp.server import do_check, do_diff
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _imported():
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
        CliRunner().invoke(main, ['build', '--project', str(p)])
        yield tmp_path, p


@test
def diff_tool_returns_empty_diff_for_built_project():
    with _imported() as (_tmp, p):
        out = do_diff(project=str(p), vs=None)
        assert out['is_empty'] is True


@test
def diff_tool_against_explicit_xlsx():
    with _imported() as (tmp_path, p):
        other = tmp_path / 'other.xlsx'
        write_simple_xlsx(other)
        out = do_diff(project=str(p), vs=f'xlsx:{other}')
        assert out['is_empty'] is True


@test
def diff_tool_reports_changed_cell():
    with _imported() as (_tmp, p):
        md = p / 'sheets' / '01_inputs.md'
        md.write_text(md.read_text().replace('0.04', '0.10'))
        out = do_diff(project=str(p), vs=None)
        assert out['is_empty'] is False
        assert any(
            sd['name'] == 'Inputs'
            and 'B1' in [c['addr'] for c in sd['cells_changed']]
            for sd in out['structured']['sheets_changed']
        )


@test
def check_tool_clean_workbook_returns_empty_list():
    with _imported() as (_tmp, p):
        out = do_check(project=str(p))
        assert out == {'issues': []}


@test
def check_tool_reports_orphaned_sheet():
    with _imported() as (_tmp, p):
        (p / 'sheets' / '99_orphan.md').write_text(
            '| (cell) | A |\n| --- | --- |\n'
        )
        out = do_check(project=str(p))
        kinds = [i['kind'] for i in out['issues']]
        assert 'sheet_file_missing_from_manifest' in kinds


@test
def diff_tool_raises_typed_error_for_missing_project():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        with catch_exceptions() as excs:
            do_diff(project=str(tmp_path / 'no-such'), vs=None)
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code == 'project_not_found'
