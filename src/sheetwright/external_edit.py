"""Detect whether the most-recent built xlsx has been edited
externally (i.e., between sheetwright sessions).
"""

from __future__ import annotations

import click

from sheetwright.build_hash import read_build_hash
from sheetwright.calc.cache import hash_xlsx
from sheetwright.project import Project


def warn_if_externally_edited(project: Project) -> None:
    """Emit a Click warning when the built xlsx no longer matches the
    recorded build hash. No-op when there's no record yet, or no
    built xlsx, or hashes match.
    """
    rec = read_build_hash(project.build_hash_path)
    if rec is None:
        return
    cfg = project.config

    if rec.name != cfg.name:
        click.echo(
            f'NOTE: recorded build was for workbook {rec.name!r}; '
            f'current workbook is {cfg.name!r}. The previous build at '
            f'build/{rec.name}.xlsx is orphaned. '
            f'Run `sheetwright build` to refresh the hash record.',
            err=True,
        )
        return

    built = project.build_dir / f'{cfg.name}.xlsx'
    if not built.is_file():
        return
    current = hash_xlsx(built)
    if current == rec.sha256:
        return
    click.echo(
        f'WARNING: {built} has been modified externally. '
        f'Run `sheetwright import {built}` to review changes.',
        err=True,
    )
