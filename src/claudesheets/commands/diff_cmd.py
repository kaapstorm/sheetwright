"""Implementation of `claudesheets diff`."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import click

from claudesheets.config import load_project
from claudesheets.diff import diff_workbooks
from claudesheets.diff.format import render
from claudesheets.exceptions import ProjectError
from claudesheets.model.workbook import Workbook
from claudesheets.project import Project
from claudesheets.source.reader import read_source
from claudesheets.xlsx.reader import read_xlsx


def run(*, project_path: str, vs: Optional[str]) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    source_wb = read_source(project.root)
    target_wb = _load_target(project, vs)

    d = diff_workbooks(target_wb, source_wb)
    click.echo(render(d))
    if not d.is_empty():
        raise click.exceptions.Exit(1)


def _load_target(project: Project, vs: Optional[str]) -> Workbook:
    if vs is None:
        cfg = load_project(project.claudesheets_toml.read_text())
        built = project.build_dir / f'{cfg.name}.xlsx'
        if not built.is_file():
            raise click.ClickException(
                f'No built xlsx at {built}. Run `claudesheets build` '
                f'first, or pass --vs.'
            )
        return read_xlsx(built)
    if vs.startswith('xlsx:'):
        return read_xlsx(Path(vs[len('xlsx:') :]))
    if vs.startswith('source:'):
        return read_source(Path(vs[len('source:') :]))
    raise click.ClickException(
        f'Invalid --vs target: {vs!r}. Expected '
        f'"xlsx:<path>" or "source:<path>".'
    )
