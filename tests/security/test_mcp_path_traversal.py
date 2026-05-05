"""Integration tests: MCP tools enforce path-containment."""

from __future__ import annotations

import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import catch_exceptions, test

from sheetwright.mcp.errors import MCPError
from sheetwright.mcp.server import (
    do_build,
    do_diff,
    do_init,
    do_reimport_stage,
    do_test,
)
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _imported_project():
    """Build a small imported+built project for integration testing."""
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
        (p / 'tests').mkdir()
        from click.testing import CliRunner

        from sheetwright.cli import main

        CliRunner().invoke(main, ['import', str(src), '--project', str(p)])
        CliRunner().invoke(main, ['build', '--project', str(p)])
        yield tmp_path, p


# ---------------------------------------------------------------------------
# do_build — out_path containment
# ---------------------------------------------------------------------------


@test
def build_out_path_escape_raises_path_outside_project():
    with _imported_project() as (_tmp, p):
        with catch_exceptions() as excs:
            do_build(project=str(p), out_path='../escape.xlsx')
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code == 'path_outside_project'


@test
def build_out_path_inside_project_succeeds():
    with _imported_project() as (_tmp, p):
        # Use the default out_path (None) — passes containment, no extra path arg.
        out = do_build(project=str(p), out_path=None)
        assert out['ok'] is True


# ---------------------------------------------------------------------------
# do_test — targets containment (uses tests_dir, not root)
# ---------------------------------------------------------------------------


@test
def test_targets_escape_raises_path_outside_project():
    with _imported_project() as (_tmp, p):
        with catch_exceptions() as excs:
            do_test(project=str(p), targets=['../../etc/x.py'])
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code == 'path_outside_project'


@test
def test_targets_inside_tests_dir_passes_containment_check():
    """Path-containment passes even if the file doesn't exist (or has no tests)."""
    with _imported_project() as (_tmp, p):
        # The file doesn't exist, so the call fails with a click_error about
        # 'no such target', but NOT a path_outside_project error.
        with catch_exceptions() as excs:
            do_test(project=str(p), targets=['tests/test_real.py'])
        if excs:
            assert isinstance(excs[0], MCPError)
            assert excs[0].code != 'path_outside_project'


# ---------------------------------------------------------------------------
# do_diff — vs path containment
# ---------------------------------------------------------------------------


@test
def diff_vs_xlsx_outside_project_raises_path_outside_project():
    with _imported_project() as (_tmp, p):
        with catch_exceptions() as excs:
            do_diff(project=str(p), vs='xlsx:/etc/passwd.xlsx')
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code == 'path_outside_project'


@test
def diff_vs_xlsx_inside_project_passes_containment():
    """Containment passes; the actual file-not-found is a different error."""
    with _imported_project() as (_tmp, p):
        with catch_exceptions() as excs:
            do_diff(project=str(p), vs=f'xlsx:{p}/build/in.xlsx')
        if excs:
            assert isinstance(excs[0], MCPError)
            assert excs[0].code != 'path_outside_project'


@test
def diff_vs_source_outside_project_raises_path_outside_project():
    with _imported_project() as (_tmp, p):
        with catch_exceptions() as excs:
            do_diff(project=str(p), vs='source:/etc')
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code == 'path_outside_project'


# ---------------------------------------------------------------------------
# do_init carve-out — its own path validation
# ---------------------------------------------------------------------------


@test
def init_dotdot_path_raises_path_outside_project():
    with catch_exceptions() as excs:
        do_init(path='../escape')
    assert excs and isinstance(excs[0], MCPError)
    assert excs[0].code == 'path_outside_project'


@test
def init_non_empty_dir_raises_path_not_empty():
    with tempfile.TemporaryDirectory() as td:
        target = Path(td)
        (target / 'file.txt').write_text('hi')
        with catch_exceptions() as excs:
            do_init(path=str(target))
        assert excs and isinstance(excs[0], MCPError)
        assert excs[0].code == 'path_not_empty'


@test
def init_fresh_absolute_dir_is_accepted():
    with tempfile.TemporaryDirectory() as td:
        target = Path(td) / 'fresh'
        out = do_init(path=str(target))
        assert out['ok'] is True


# ---------------------------------------------------------------------------
# do_reimport_stage.xlsx carve-out — no path_outside_project
# ---------------------------------------------------------------------------


@test
def reimport_stage_xlsx_outside_project_does_not_raise_path_outside_project():
    """xlsx arg is a documented carve-out; the error must NOT be path_outside_project."""
    with _imported_project() as (_tmp, p):
        outside_xlsx = Path('/nonexistent_outside_path/file.xlsx')
        with catch_exceptions() as excs:
            do_reimport_stage(
                xlsx=str(outside_xlsx), project=str(p), force=False
            )
        # The call will fail (file doesn't exist), but NOT with path_outside_project.
        # The error may be an MCPError with a different code, or a raw exception —
        # either way it must not be path_outside_project.
        assert excs
        if isinstance(excs[0], MCPError):
            assert excs[0].code != 'path_outside_project'
