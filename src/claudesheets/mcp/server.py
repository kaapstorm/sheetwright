"""MCP server for claudesheets.

Exposes each CLI command as an MCP tool. The server is a thin facade
over `claudesheets.commands.*.run` and the pure helpers in
`claudesheets.diff` / `reimport`.
"""

from __future__ import annotations

from typing import Any, Dict, Optional

import click
from mcp.server.fastmcp import FastMCP

from claudesheets.diff import diff_workbooks
from claudesheets.diff.check import check_workbook
from claudesheets.diff.loaders import load_target
from claudesheets.exceptions import ProjectError
from claudesheets.mcp.errors import MCPError, classify_click_error
from claudesheets.mcp.shaping import check_issues_to_dicts, diff_to_dict
from claudesheets.project import Project
from claudesheets.source.reader import read_source


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


def build_server() -> FastMCP:
    """Construct and return the claudesheets MCP server."""
    mcp = FastMCP('claudesheets')
    mcp.tool()(do_ping)
    mcp.tool()(do_diff)
    mcp.tool()(do_check)
    return mcp
