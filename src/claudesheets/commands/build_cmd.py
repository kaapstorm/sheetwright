"""Implementation of `claudesheets build`."""

from __future__ import annotations

from pathlib import Path

import click

from claudesheets.bulk import build_bulk_cache
from claudesheets.config import load_project
from claudesheets.exceptions import ProjectError
from claudesheets.project import Project
from claudesheets.source.reader import read_source
from claudesheets.xlsx.writer import write_xlsx


def run(*, project_path: str, out_path: str | None) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    cfg = load_project(project.claudesheets_toml.read_text())

    wb = read_source(project.root)
    build_bulk_cache(project.root)

    project.build_dir.mkdir(parents=True, exist_ok=True)
    target = (
        Path(out_path) if out_path else project.build_dir / f'{cfg.name}.xlsx'
    )
    write_xlsx(wb, target)

    click.echo(f'Built {target}')
