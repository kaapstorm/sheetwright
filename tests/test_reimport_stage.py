from __future__ import annotations

import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import catch_exceptions, test

from sheetwright.project import Project
from sheetwright.reimport.flow import StagedReimport, stage_reimport
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _populated():
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
        from click.testing import CliRunner

        from sheetwright.cli import main

        CliRunner().invoke(main, ['import', str(src), '--project', str(p)])
        yield tmp_path, p, src


@test
def stage_returns_empty_diff_when_xlsx_unchanged():
    with _populated() as (_tmp, p, src):
        project = Project.open(p)
        staged = stage_reimport(project, src, flatten=False, force=False)
        assert isinstance(staged, StagedReimport)
        assert staged.diff.is_empty()
        assert 'no changes' in staged.rendered_diff.lower()
        assert staged.xlsx_path == src.resolve()


@test
def stage_returns_diff_for_changed_xlsx():
    with _populated() as (tmp_path, p, _src):
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


@test
def stage_raises_on_external_refs_without_flatten():
    import click

    from tests.fixtures.external_xlsx import write_xlsx_with_external_ref

    with _populated() as (_tmp, p, _src):
        bad_src = p / 'bad.xlsx'
        write_xlsx_with_external_ref(bad_src, cached_value=42.0)

        project = Project.open(p)
        with catch_exceptions() as excs:
            stage_reimport(project, bad_src, flatten=False, force=False)
        assert excs and isinstance(excs[0], click.ClickException)
        assert 'external reference' in str(excs[0])
