"""Implementation of `claudesheets mcp`."""

from __future__ import annotations


def run() -> None:
    from claudesheets.mcp import build_server

    server = build_server()
    server.run(transport='stdio')
