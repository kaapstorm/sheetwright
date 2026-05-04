"""Tests for mutating MCP tools: do_init, do_import_xlsx, do_build."""

from pathlib import Path

import pytest
from unmagic import fixture, use

from claudesheets.mcp.errors import MCPError
from claudesheets.mcp.server import do_build, do_import_xlsx, do_init
from tests.fixtures.workbooks import write_simple_xlsx


def test_init_tool_creates_skeleton(tmp_path: Path):
    target = tmp_path / 'fresh'
    out = do_init(path=str(target))
    assert out['ok'] is True
    assert (target / 'claudesheets.toml').is_file()


def test_init_tool_refuses_non_empty_dir(tmp_path: Path):
    target = tmp_path / 'fresh'
    target.mkdir()
    (target / 'file.txt').write_text('hi')
    with pytest.raises(MCPError) as excinfo:
        do_init(path=str(target))
    assert excinfo.value.code


@fixture
def empty_project(tmp_path: Path):
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "in"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    yield p


@use(empty_project)
def test_import_xlsx_tool(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    out = do_import_xlsx(
        xlsx=str(src),
        project=str(empty_project()),
        archive=False,
        flatten=False,
    )
    assert out['ok'] is True
    assert (empty_project() / 'sheets' / '01_inputs.md').is_file()


@use(empty_project)
def test_build_tool(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    p = empty_project()
    do_import_xlsx(xlsx=str(src), project=str(p), archive=False, flatten=False)
    out = do_build(project=str(p), out_path=None)
    assert out['ok'] is True
    assert (p / 'build' / 'in.xlsx').is_file()


@use(empty_project)
def test_import_xlsx_rejects_when_source_already_populated(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    p = empty_project()
    do_import_xlsx(xlsx=str(src), project=str(p), archive=False, flatten=False)
    with pytest.raises(MCPError) as excinfo:
        do_import_xlsx(
            xlsx=str(src), project=str(p), archive=False, flatten=False
        )
    assert excinfo.value.code == 'reimport_required'
