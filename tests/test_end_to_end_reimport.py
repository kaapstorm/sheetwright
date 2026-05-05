import json
import tempfile
from contextlib import contextmanager
from pathlib import Path

import openpyxl
from click.testing import CliRunner
from testsweet import test

from sheetwright.calc.cache import hash_xlsx
from sheetwright.cli import main
from sheetwright.exceptions import StaleSessionFormatError
from sheetwright.project import Project
from sheetwright.reimport.flow import apply_session
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


@contextmanager
def _staged_project():
    """A project with initial source and a second xlsx staged for re-import.

    Does not require LibreOffice — uses CLI import only.
    """
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'v1.xlsx'
        write_simple_xlsx(src)

        p = tmp_path / 'proj'
        p.mkdir()
        (p / 'sheetwright.toml').write_text(
            '[project]\nname = "proj"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "proj"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()

        runner = CliRunner()
        # initial import to populate source
        runner.invoke(main, ['import', str(src), '--project', str(p)])

        # create a slightly different v2 xlsx
        v2 = tmp_path / 'v2.xlsx'
        wb = openpyxl.load_workbook(src)
        wb['Inputs']['B1'] = 0.10
        wb.save(v2)

        yield p, src, v2, runner


@test
@requires_libreoffice
def full_escape_hatch_loop():
    with _baseline() as (p, _orig_src):
        runner = CliRunner()

        edited = p / 'build' / 'in.xlsx'
        import openpyxl as _openpyxl

        wb = _openpyxl.load_workbook(edited)
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


@test
def staging_copy_invariant():
    """stage_reimport copies the xlsx into .sheetwright/staged/<sha>.xlsx."""
    with _staged_project() as (p, _src, v2, runner):
        sha = hash_xlsx(v2)
        runner.invoke(main, ['import', str(v2), '-I', '--project', str(p)])
        staging_copy = p / '.sheetwright' / 'staged' / f'{sha}.xlsx'
        assert staging_copy.is_file(), 'Staging copy was not created'
        assert hash_xlsx(staging_copy) == sha


@test
def apply_uses_staging_copy_when_original_modified():
    """apply_session succeeds even if the original xlsx is modified post-stage."""
    with _staged_project() as (p, _src, v2, runner):
        r = runner.invoke(main, ['import', str(v2), '-I', '--project', str(p)])
        assert r.exit_code == 0, r.output

        # corrupt the original xlsx after staging
        v2.write_bytes(b'garbage bytes that corrupt the xlsx')

        r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
        assert r.exit_code == 0, r.output


@test
def apply_rejects_when_staging_copy_is_gone():
    """apply raises a clear error when the staging copy has been removed."""
    with _staged_project() as (p, _src, v2, runner):
        r = runner.invoke(main, ['import', str(v2), '-I', '--project', str(p)])
        assert r.exit_code == 0, r.output

        sha = hash_xlsx(v2)
        staging_copy = p / '.sheetwright' / 'staged' / f'{sha}.xlsx'
        staging_copy.unlink()

        r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
        assert r.exit_code != 0
        assert 'no longer exists' in r.output.lower()


@test
def apply_raises_stale_session_format_error():
    """apply raises StaleSessionFormatError for a pre-staging-copy session."""
    with _staged_project() as (p, _src, v2, runner):
        # Write an old-format session (no staged_filename)
        session_path = p / '.sheetwright' / 'reimport.json'
        session_path.parent.mkdir(parents=True, exist_ok=True)
        old_data = {
            'xlsx_path': str(v2),
            'xlsx_sha256': hash_xlsx(v2),
            'diff_summary': 'n/a',
            'created_at': '2026-01-01T00:00:00+00:00',
        }
        session_path.write_text(json.dumps(old_data))

        project = Project.open(p)
        try:
            apply_session(project, archive=False, flatten=False)
            assert False, 'Expected StaleSessionFormatError'
        except StaleSessionFormatError as e:
            assert 're-stage' in str(e).lower()


@test
def archive_sources_from_staging_copy():
    """Archived xlsx matches the staging copy even if original was modified."""
    with _staged_project() as (p, _src, v2, runner):
        sha = hash_xlsx(v2)

        r = runner.invoke(main, ['import', str(v2), '-I', '--project', str(p)])
        assert r.exit_code == 0, r.output

        staging_copy = p / '.sheetwright' / 'staged' / f'{sha}.xlsx'
        staging_bytes = staging_copy.read_bytes()

        # modify original after staging
        v2.write_bytes(b'modified')

        r = runner.invoke(
            main, ['import', '--apply', '--archive', '--project', str(p)]
        )
        assert r.exit_code == 0, r.output

        imports_dir = p / 'imports'
        archived = sorted(imports_dir.glob('*.xlsx'))
        assert len(archived) == 1, 'Expected exactly one archived file'
        assert archived[0].read_bytes() == staging_bytes
