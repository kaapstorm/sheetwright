"""Implementation of `sheetwright recalc`."""

from __future__ import annotations

from pathlib import Path

import click

from sheetwright.calc import get_calc_engine
from sheetwright.calc.cache import (
    hash_xlsx,
    read_cached,
    write_cached,
)
from sheetwright.exceptions import ProjectError
from sheetwright.project import Project
from sheetwright.security import SecurityLimits, get_operator_limits


CACHE_HIT_MESSAGE = 'cache hit'


def run(*, project_path: str, force: bool) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    from sheetwright.external_edit import warn_if_externally_edited

    warn_if_externally_edited(project)

    cfg = project.config
    built = project.build_dir / f'{cfg.name}.xlsx'
    if not built.is_file():
        raise click.ClickException(
            f'No built xlsx at {built}. Run `sheetwright build` first.'
        )

    key = hash_xlsx(built)
    if not force:
        cached = read_cached(project.calc_cache_dir, key)
        if cached is not None:
            click.echo(f'{CACHE_HIT_MESSAGE}: {key[:12]}')
            return

    limits = SecurityLimits.effective(
        get_operator_limits(), project.config.security
    )
    engine = get_calc_engine(cfg.calc_engine)
    result = engine.evaluate(built, limits=limits)
    path = write_cached(project.calc_cache_dir, key, result)
    click.echo(f'recalculated: {path}')
