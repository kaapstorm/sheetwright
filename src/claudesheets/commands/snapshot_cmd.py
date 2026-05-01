"""Implementation of `claudesheets snapshot`."""

from __future__ import annotations

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
from claudesheets.snapshot import (
    Snapshot,
    diff_snapshots,
    snapshot_from_calc_result,
)
from claudesheets.source.reader import read_source


def run(*, project_path: str, update: bool) -> None:
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
    cached = read_cached(project.calc_cache_dir, key)
    if cached is None:
        cached = get_calc_engine(cfg.calc_engine).evaluate(built)
        write_cached(project.calc_cache_dir, key, cached)

    workbook = read_source(project.root)
    current = snapshot_from_calc_result(cached, workbook)
    snap_path = project.snapshots_dir / f'{cfg.name}.json'

    if not snap_path.is_file():
        current.write(snap_path)
        click.echo(f'initialized snapshot at {snap_path}')
        return

    if update:
        current.write(snap_path)
        click.echo(f'updated snapshot at {snap_path}')
        return

    saved = Snapshot.read(snap_path)
    diffs = diff_snapshots(saved, current)
    if not diffs:
        click.echo('no changes')
        return

    for sheet, addr, old, new in diffs:
        click.echo(f'  {sheet}!{addr}: {old!r} -> {new!r}')
    raise click.exceptions.Exit(1)
