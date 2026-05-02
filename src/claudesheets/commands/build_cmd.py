"""Implementation of `claudesheets build`."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

import click

from claudesheets.build_hash import BuildHashRecord, write_build_hash
from claudesheets.bulk import build_bulk_cache
from claudesheets.calc.cache import hash_xlsx
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
    write_build_hash(
        project.build_hash_path,
        BuildHashRecord(
            name=cfg.name,
            sha256=hash_xlsx(target),
            built_at=datetime.now(timezone.utc).isoformat(),
        ),
    )

    click.echo(f'Built {target}')
