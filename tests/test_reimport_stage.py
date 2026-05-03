from __future__ import annotations

from pathlib import Path

from unmagic import fixture, use

from claudesheets.project import Project
from claudesheets.reimport.flow import StagedReimport, stage_reimport
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def populated(tmp_path: Path):
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
    yield p, src


@use(populated)
def test_stage_returns_empty_diff_when_xlsx_unchanged():
    p, src = populated()
    project = Project.open(p)
    staged = stage_reimport(project, src, flatten=False, force=False)
    assert isinstance(staged, StagedReimport)
    assert staged.diff.is_empty()
    assert 'no changes' in staged.rendered_diff.lower()
    assert staged.xlsx_path == src.resolve()


@use(populated)
def test_stage_returns_diff_for_changed_xlsx(tmp_path: Path):
    p, _src = populated()
    project = Project.open(p)

    new_src = tmp_path / 'changed.xlsx'
    import openpyxl

    wb = openpyxl.Workbook()
    s = wb.active
    s.title = 'Inputs'
    s['A1'], s['B1'] = 'growth_rate', 0.99
    s['A2'], s['B2'] = 'base_revenue', 1_000_000
    wb.create_sheet('Outputs')
    wb.save(new_src)

    staged = stage_reimport(project, new_src, flatten=False, force=False)
    assert not staged.diff.is_empty()
    assert 'Inputs!B1' in staged.rendered_diff


@use(populated)
def test_stage_raises_on_external_refs_without_flatten():
    """stage_reimport raises ClickException on external refs (no flatten).

    The CLI catches and reformats; MCP catches and translates to a
    typed error.
    """
    import click
    import pytest

    from tests.fixtures.external_xlsx import write_xlsx_with_external_ref

    p, _ = populated()
    bad_src = p / 'bad.xlsx'
    write_xlsx_with_external_ref(bad_src, cached_value=42.0)

    project = Project.open(p)
    with pytest.raises(click.ClickException, match='external reference'):
        stage_reimport(project, bad_src, flatten=False, force=False)
