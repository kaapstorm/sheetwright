"""MCP server for sheetwright.

Exposes each CLI command as an MCP tool. The server is a thin facade
over `sheetwright.commands.*.run` and the pure helpers in
`sheetwright.diff` / `reimport`.
"""

from __future__ import annotations

import io
from contextlib import redirect_stdout
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import click
from mcp.server.fastmcp import FastMCP

from sheetwright.commands.build_cmd import run as _build_run
from sheetwright.commands.import_cmd import run as _import_run
from sheetwright.commands.init_cmd import run as _init_run
from sheetwright.commands.recalc_cmd import run as _recalc_run
from sheetwright.commands.snapshot_cmd import run as _snapshot_run
from sheetwright.commands.test_cmd import run as _test_run
from sheetwright.diff import diff_workbooks
from sheetwright.diff.check import check_workbook
from sheetwright.diff.loaders import load_target
from sheetwright.exceptions import ProjectError
from sheetwright.mcp.errors import MCPError, classify_click_error
from sheetwright.mcp.shaping import check_issues_to_dicts, diff_to_dict
from sheetwright.project import Project
from sheetwright.reimport import (
    ReimportSession,
    apply_session,
    clear_session,
    save_session,
    stage_reimport,
)
from sheetwright.source.reader import read_source


def _ok(message: str, **extra: Any) -> Dict[str, Any]:
    return {'ok': True, 'message': message, **extra}


def do_ping() -> str:
    """Smoke test — returns 'pong'."""
    return 'pong'


def _open_project(project: str) -> Project:
    """Open a Project, translating ProjectError to MCPError."""
    try:
        return Project.open(project)
    except ProjectError as e:
        raise MCPError('project_not_found', str(e))


def do_diff(project: str, vs: Optional[str] = None) -> Dict[str, Any]:
    """Diff the project's source against a target workbook.

    `vs`:
    - `None` — compare to `build/<name>.xlsx`.
    - `'xlsx:<path>'` — compare to a specific xlsx.
    - `'source:<path>'` — compare to another project's source dir.

    Returns `{is_empty, rendered, structured}`.
    """
    proj = _open_project(project)
    source_wb = read_source(proj.root)
    try:
        target_wb = load_target(proj, vs)
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)
    return diff_to_dict(diff_workbooks(target_wb, source_wb))


def do_check(project: str) -> Dict[str, Any]:
    """Lint dangling refs, missing names, schema mismatches.

    Returns `{issues: [{kind, detail, location}, ...]}`.
    """
    proj = _open_project(project)
    wb = read_source(proj.root)
    issues = check_workbook(wb, proj)
    return {'issues': check_issues_to_dicts(issues)}


def _run_capturing(fn: Any, **kwargs: Any) -> str:
    """Run `fn(**kwargs)` with stdout captured; return captured text.

    Translates `ClickException` to `MCPError`. Used by tools that
    delegate to a `commands/*.run()` function and don't expect
    `click.exceptions.Exit` (which signals "result with diffs").
    """
    buf = io.StringIO()
    try:
        with redirect_stdout(buf):
            fn(**kwargs)
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)
    return buf.getvalue()


def do_init(path: str) -> Dict[str, Any]:
    """Scaffold an empty sheetwright project at `path`."""
    captured = _run_capturing(_init_run, path=path).strip()
    return _ok(captured or f'Initialised sheetwright project at {path}')


def do_import_xlsx(
    xlsx: str,
    project: str,
    archive: bool = False,
    flatten: bool = False,
) -> Dict[str, Any]:
    """Read an .xlsx file into source form (initial import only).

    Re-import is NOT exposed here; use the re-import tools
    (`do_reimport_stage` / `do_reimport_apply` / `do_reimport_abort`).
    """
    proj = _open_project(project)
    if proj.has_source():
        raise MCPError(
            'reimport_required',
            'sheets/ is non-empty; use do_reimport_stage / '
            'do_reimport_apply for the merge flow.',
        )
    captured = _run_capturing(
        _import_run,
        xlsx_path=xlsx,
        project_path=project,
        archive=archive,
        flatten=flatten,
        non_interactive=False,
        apply=False,
        abort=False,
        force=False,
    ).strip()
    return _ok(captured or f'Imported {xlsx} into {project}')


def do_build(project: str, out_path: Optional[str] = None) -> Dict[str, Any]:
    """Compile sources into an .xlsx."""
    captured = _run_capturing(
        _build_run, project_path=project, out_path=out_path
    ).strip()
    return _ok(captured or 'build complete')


def do_recalc(project: str, force: bool = False) -> Dict[str, Any]:
    """Run the calc engine; cache results."""
    captured = _run_capturing(
        _recalc_run, project_path=project, force=force
    ).strip()
    return _ok(captured or 'recalc complete')


def do_snapshot(project: str, update: bool = False) -> Dict[str, Any]:
    """Compare or update the snapshot of calculated values.

    Returns {ok, message, has_diffs}. `has_diffs=True` means the
    saved snapshot differs from the current calculation; this is a
    successful tool result, not a failure.
    """
    buf = io.StringIO()
    has_diffs = False
    try:
        with redirect_stdout(buf):
            _snapshot_run(project_path=project, update=update)
    except click.exceptions.Exit as e:
        has_diffs = e.exit_code != 0
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)
    return {
        'ok': True,
        'message': buf.getvalue().strip(),
        'has_diffs': has_diffs,
    }


def do_test(
    project: str, targets: Optional[List[str]] = None
) -> Dict[str, Any]:
    """Run the project's testsweet tests.

    Returns `{passed: bool, output: str}`.
    """
    buf = io.StringIO()
    passed = True
    try:
        with redirect_stdout(buf):
            _test_run(project_path=project, targets=targets or [])
    except click.exceptions.Exit as e:
        passed = e.exit_code == 0
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)
    return {'passed': passed, 'output': buf.getvalue()}


def do_reimport_stage(
    xlsx: str,
    project: str,
    flatten: bool = False,
    force: bool = False,
) -> Dict[str, Any]:
    """Compute and stage a re-import diff.

    If the diff is non-empty, saves a session on disk that
    `do_reimport_apply` can later commit. Returns the diff plus
    `xlsx_path` and `xlsx_sha256` for cross-call verification.
    """
    proj = _open_project(project)
    try:
        staged = stage_reimport(proj, Path(xlsx), flatten=flatten, force=force)
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)

    out = diff_to_dict(staged.diff)
    out['rendered_diff'] = staged.rendered_diff
    out['xlsx_path'] = str(staged.xlsx_path)
    out['xlsx_sha256'] = staged.xlsx_sha256

    if not staged.diff.is_empty():
        save_session(
            proj.reimport_session_path,
            ReimportSession(
                xlsx_path=str(staged.xlsx_path),
                xlsx_sha256=staged.xlsx_sha256,
                diff_summary=staged.rendered_diff,
                created_at=datetime.now(timezone.utc).isoformat(),
            ),
        )

    return out


def do_reimport_apply(project: str, archive: bool = False) -> Dict[str, Any]:
    """Complete a previously staged re-import."""
    proj = _open_project(project)
    buf = io.StringIO()
    try:
        with redirect_stdout(buf):
            # flatten was decided at stage time and reflected in the
            # saved session; apply just writes what was staged.
            apply_session(proj, archive=archive, flatten=False)
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)
    return _ok(buf.getvalue().strip() or 'Staged re-import applied.')


def do_reimport_abort(project: str) -> Dict[str, Any]:
    """Discard a previously staged re-import session."""
    proj = _open_project(project)
    clear_session(proj.reimport_session_path)
    return _ok('Re-import session cleared.')


def build_server() -> FastMCP:
    """Construct and return the sheetwright MCP server."""
    mcp = FastMCP('sheetwright')
    mcp.tool()(do_ping)
    mcp.tool()(do_diff)
    mcp.tool()(do_check)
    mcp.tool()(do_init)
    mcp.tool()(do_import_xlsx)
    mcp.tool()(do_build)
    mcp.tool()(do_recalc)
    mcp.tool()(do_snapshot)
    mcp.tool()(do_test)
    mcp.tool()(do_reimport_stage)
    mcp.tool()(do_reimport_apply)
    mcp.tool()(do_reimport_abort)
    return mcp
