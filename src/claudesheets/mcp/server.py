"""MCP server for claudesheets.

Exposes each CLI command as an MCP tool. The server is a thin facade
over `claudesheets.commands.*.run` and the pure helpers in
`claudesheets.diff` / `reimport`.
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP


def do_ping() -> str:
    """Smoke test — returns 'pong'."""
    return 'pong'


def build_server() -> FastMCP:
    """Construct and return the claudesheets MCP server."""
    mcp = FastMCP('claudesheets')
    mcp.tool()(do_ping)
    return mcp
