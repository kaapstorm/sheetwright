"""Structured errors raised by MCP tools."""

from __future__ import annotations

import click


class MCPError(Exception):
    """A typed error returned to MCP clients.

    `code` is a stable machine-readable string (e.g. 'project_not_found',
    'external_refs', 'reimport_required', 'no_staged_session'). Clients
    branch on `code`; humans read `message`.
    """

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


def classify_click_error(e: click.ClickException) -> str:
    """Map known ClickException messages to a stable error code.

    Unknown messages fall back to 'click_error'. Add codes here as
    real clients prove they want to branch on them.
    """
    msg = e.message
    lower = msg.lower()
    if 'external reference' in lower:
        return 'external_refs'
    if 'uncommitted' in lower:
        return 'uncommitted_source'
    if 'not a claudesheets project' in lower:
        return 'project_not_found'
    if 'no built xlsx' in lower:
        return 'build_missing'
    if 'no staged re-import session' in lower:
        return 'no_staged_session'
    if 'staged xlsx' in lower:
        return 'staged_xlsx_changed'
    return 'click_error'
