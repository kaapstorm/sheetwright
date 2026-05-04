"""MCP server for claudesheets.

Exposes each CLI command as an MCP tool. The server is a thin facade
over `claudesheets.commands.*.run` and the pure helpers in
`claudesheets.diff` / `reimport`.
"""

from __future__ import annotations

import io
from contextlib import redirect_stdout
from typing import Any, Callable, Dict, Optional, TypeVar

import click
from mcp.server.fastmcp import FastMCP

from claudesheets.commands.build_cmd import run as _build_run
from claudesheets.commands.import_cmd import run as _import_run
from claudesheets.commands.init_cmd import run as _init_run
from claudesheets.diff import diff_workbooks
from claudesheets.diff.check import check_workbook
from claudesheets.diff.loaders import load_target
from claudesheets.exceptions import ProjectError
from claudesheets.mcp.errors import MCPError, classify_click_error
from claudesheets.mcp.shaping import check_issues_to_dicts, diff_to_dict
from claudesheets.project import Project
from claudesheets.source.reader import read_source


T = TypeVar('T')


def _capture(fn: Callable[..., T], *args: Any, **kwargs: Any) -> tuple[T, str]:
    """Run `fn`; return its result + anything click.echo'd to stdout."""
    buf = io.StringIO()
    with redirect_stdout(buf):
        result = fn(*args, **kwargs)
    return result, buf.getvalue()


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


def do_init(path: str) -> Dict[str, Any]:
    """Scaffold an empty claudesheets project at `path`."""
    try:
        _, captured = _capture(_init_run, path)
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)
    msg = captured.strip() or f'Initialised claudesheets project at {path}'
    return _ok(msg)


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
    if any(proj.sheets_dir.iterdir()):
        raise MCPError(
            'reimport_required',
            'sheets/ is non-empty; use do_reimport_stage / '
            'do_reimport_apply for the merge flow.',
        )
    try:
        _, captured = _capture(
            _import_run,
            xlsx_path=xlsx,
            project_path=project,
            archive=archive,
            flatten=flatten,
            non_interactive=False,
            apply=False,
            abort=False,
            force=False,
        )
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)
    return _ok(captured.strip() or f'Imported {xlsx} into {project}')


def do_build(project: str, out_path: Optional[str] = None) -> Dict[str, Any]:
    """Compile sources into an .xlsx."""
    try:
        _, captured = _capture(
            _build_run, project_path=project, out_path=out_path
        )
    except click.ClickException as e:
        raise MCPError(classify_click_error(e), e.message)
    return _ok(captured.strip() or 'build complete')


def build_server() -> FastMCP:
    """Construct and return the claudesheets MCP server."""
    mcp = FastMCP('claudesheets')
    mcp.tool()(do_ping)
    mcp.tool()(do_diff)
    mcp.tool()(do_check)
    mcp.tool()(do_init)
    mcp.tool()(do_import_xlsx)
    mcp.tool()(do_build)
    return mcp
