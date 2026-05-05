"""Loaders that produce a Workbook from various sources.

Used by `sheetwright diff` and the MCP `do_diff` tool.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import click

from sheetwright.model.workbook import Workbook
from sheetwright.project import Project
from sheetwright.security import SecurityLimits, get_operator_limits
from sheetwright.source.reader import read_source
from sheetwright.xlsx.reader import read_xlsx


def load_target(project: Project, vs: Optional[str]) -> Workbook:
    """Resolve the diff target.

    `vs` is `None` (compare to `build/<name>.xlsx`),
    `'xlsx:<path>'`, or `'source:<path>'`.

    Raises `click.ClickException` for caller-facing errors.
    """
    limits = SecurityLimits.effective(
        get_operator_limits(), project.config.security
    )
    if vs is None:
        built = project.build_dir / f'{project.config.name}.xlsx'
        if not built.is_file():
            raise click.ClickException(
                f'No built xlsx at {built}. Run `sheetwright build` '
                f'first, or pass --vs.'
            )
        return read_xlsx(built, limits=limits)
    if vs.startswith('xlsx:'):
        return read_xlsx(Path(vs[len('xlsx:') :]), limits=limits)
    if vs.startswith('source:'):
        return read_source(Path(vs[len('source:') :]))
    raise click.ClickException(
        f'Invalid --vs target: {vs!r}. Expected '
        f'"xlsx:<path>" or "source:<path>".'
    )
