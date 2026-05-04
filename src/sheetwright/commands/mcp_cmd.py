"""Implementation of `sheetwright mcp`."""

from __future__ import annotations


def run() -> None:
    from sheetwright.mcp import build_server

    server = build_server()
    server.run(transport='stdio')
