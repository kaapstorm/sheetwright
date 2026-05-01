"""Implementation of `claudesheets recalc`."""

from __future__ import annotations

from pathlib import Path

import click

from claudesheets.calc import get_calc_engine
from claudesheets.calc.cache import (
    hash_xlsx,
    read_cached,
    write_cached,
)
from claudesheets.config import load_project
from claudesheets.exceptions import ProjectError
from claudesheets.project import Project


CACHE_HIT_MESSAGE = 'cache hit'


def run(*, project_path: str, force: bool) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    cfg = load_project(project.claudesheets_toml.read_text())
    built = project.build_dir / f'{cfg.name}.xlsx'
    if not built.is_file():
        raise click.ClickException(
            f'No built xlsx at {built}. Run `claudesheets build` first.'
        )

    key = hash_xlsx(built)
    if not force:
        cached = read_cached(project.calc_cache_dir, key)
        if cached is not None:
            click.echo(f'{CACHE_HIT_MESSAGE}: {key[:12]}')
            return

    engine = get_calc_engine(cfg.calc_engine)
    result = engine.evaluate(built)
    path = write_cached(project.calc_cache_dir, key, result)
    click.echo(f'recalculated: {path}')
