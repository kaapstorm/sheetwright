"""Loaders that produce a Workbook from various sources.

Used by `sheetwright diff` and the MCP `do_diff` tool.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Optional

import click

from sheetwright.model.workbook import Workbook
from sheetwright.project import Project
from sheetwright.security import SecurityLimits, get_operator_limits
from sheetwright.source.reader import read_source
from sheetwright.xlsx.reader import read_xlsx


@dataclass(frozen=True)
class VsTarget:
    kind: Literal['xlsx', 'source']
    path: Path


def parse_vs_target(vs: str) -> VsTarget:
    """Parse a 'vs=' specifier. Raises click.ClickException on bad
    format. (Phase 2 will replace ClickException; Task 3 keeps the
    existing behaviour at the CLI seam.)
    """
    if vs.startswith('xlsx:'):
        return VsTarget(kind='xlsx', path=Path(vs.removeprefix('xlsx:')))
    if vs.startswith('source:'):
        return VsTarget(kind='source', path=Path(vs.removeprefix('source:')))
    raise click.ClickException(
        f'Invalid --vs target: {vs!r}. Expected '
        f'"xlsx:<path>" or "source:<path>".'
    )


def load_parsed_target(
    project: Project, target: Optional[VsTarget], *, limits: SecurityLimits
) -> Workbook:
    """Resolve the diff target from a pre-parsed `VsTarget`.

    `target` is `None` (compare to `build/<name>.xlsx`) or a
    `VsTarget` produced by `parse_vs_target`.

    Raises `click.ClickException` for caller-facing errors.
    """
    if target is None:
        built = project.build_dir / f'{project.config.name}.xlsx'
        if not built.is_file():
            raise click.ClickException(
                f'No built xlsx at {built}. Run `sheetwright build` '
                f'first, or pass --vs.'
            )
        return read_xlsx(built, limits=limits)
    if target.kind == 'xlsx':
        return read_xlsx(target.path, limits=limits)
    return read_source(target.path)


def load_target(project: Project, vs: Optional[str]) -> Workbook:
    """Resolve the diff target.

    `vs` is `None` (compare to `build/<name>.xlsx`),
    `'xlsx:<path>'`, or `'source:<path>'`.

    Raises `click.ClickException` for caller-facing errors.
    """
    limits = SecurityLimits.effective(
        get_operator_limits(), project.config.security
    )
    target = parse_vs_target(vs) if vs is not None else None
    return load_parsed_target(project, target, limits=limits)
