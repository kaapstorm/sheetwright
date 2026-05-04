from pathlib import Path

import pytest
from unmagic import fixture, use

from claudesheets.mcp.errors import MCPError
from claudesheets.mcp.server import do_check, do_diff
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def imported(tmp_path: Path):
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
    from click.testing import CliRunner

    from claudesheets.cli import main

    CliRunner().invoke(main, ['import', str(src), '--project', str(p)])
    CliRunner().invoke(main, ['build', '--project', str(p)])
    yield p


@use(imported)
def test_diff_tool_returns_empty_diff_for_built_project():
    p = imported()
    out = do_diff(project=str(p), vs=None)
    assert out['is_empty'] is True


@use(imported)
def test_diff_tool_against_explicit_xlsx(tmp_path: Path):
    p = imported()
    other = tmp_path / 'other.xlsx'
    write_simple_xlsx(other)
    out = do_diff(project=str(p), vs=f'xlsx:{other}')
    assert out['is_empty'] is True


@use(imported)
def test_diff_tool_reports_changed_cell():
    p = imported()
    md = p / 'sheets' / '01_inputs.md'
    md.write_text(md.read_text().replace('0.04', '0.10'))
    out = do_diff(project=str(p), vs=None)
    assert out['is_empty'] is False
    assert any(
        sd['name'] == 'Inputs'
        and 'B1' in [c['addr'] for c in sd['cells_changed']]
        for sd in out['structured']['sheets_changed']
    )


@fixture
def clean_project(tmp_path: Path):
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "x"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    yield p


@use(clean_project)
def test_check_tool_clean_workbook_returns_empty_list():
    p = clean_project()
    out = do_check(project=str(p))
    assert out == {'issues': []}


@use(clean_project)
def test_check_tool_reports_orphaned_sheet():
    p = clean_project()
    (p / 'sheets' / '99_orphan.md').write_text(
        '| (cell) | A |\n| --- | --- |\n'
    )
    out = do_check(project=str(p))
    kinds = [i['kind'] for i in out['issues']]
    assert 'sheet_file_missing_from_manifest' in kinds


def test_diff_tool_raises_typed_error_for_missing_project(tmp_path: Path):
    with pytest.raises(MCPError) as excinfo:
        do_diff(project=str(tmp_path / 'no-such'), vs=None)
    assert excinfo.value.code == 'project_not_found'
