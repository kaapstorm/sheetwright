"""Tests for mutating MCP tools: do_init, do_import_xlsx, do_build."""

import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import catch_exceptions, test

from sheetwright.mcp.errors import MCPError
from sheetwright.mcp.server import do_build, do_import_xlsx, do_init
from tests.fixtures.workbooks import write_simple_xlsx


@test
def init_tool_creates_skeleton():
    with tempfile.TemporaryDirectory() as td:
        target = Path(td) / 'fresh'
        out = do_init(path=str(target))
        assert out['ok'] is True
        assert (target / 'sheetwright.toml').is_file()


@test
def init_tool_refuses_non_empty_dir():
    with tempfile.TemporaryDirectory() as td:
        target = Path(td) / 'fresh'
        target.mkdir()
        (target / 'file.txt').write_text('hi')
        with catch_exceptions() as excs:
            do_init(path=str(target))
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code


@contextmanager
def _empty_project():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
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
        yield tmp_path, p


@test
def import_xlsx_tool():
    with _empty_project() as (tmp_path, p):
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        out = do_import_xlsx(
            xlsx=str(src),
            project=str(p),
            archive=False,
            flatten=False,
        )
        assert out['ok'] is True
        assert (p / 'sheets' / '01_inputs.md').is_file()


@test
def build_tool():
    with _empty_project() as (tmp_path, p):
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        do_import_xlsx(
            xlsx=str(src), project=str(p), archive=False, flatten=False
        )
        out = do_build(project=str(p), out_path=None)
        assert out['ok'] is True
        assert (p / 'build' / 'in.xlsx').is_file()


@test
def import_xlsx_rejects_when_source_already_populated():
    with _empty_project() as (tmp_path, p):
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        do_import_xlsx(
            xlsx=str(src), project=str(p), archive=False, flatten=False
        )
        with catch_exceptions() as excs:
            do_import_xlsx(
                xlsx=str(src), project=str(p), archive=False, flatten=False
            )
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code == 'reimport_required'
