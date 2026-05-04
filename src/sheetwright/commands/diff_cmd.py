"""Implementation of `sheetwright diff`."""

from __future__ import annotations

from typing import Optional

import click

from sheetwright.diff import diff_workbooks
from sheetwright.diff.format import render
from sheetwright.diff.loaders import load_target
from sheetwright.exceptions import ProjectError
from sheetwright.project import Project
from sheetwright.source.reader import read_source


def run(*, project_path: str, vs: Optional[str]) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    source_wb = read_source(project.root)
    target_wb = load_target(project, vs)

    d = diff_workbooks(target_wb, source_wb)
    click.echo(render(d))
    if not d.is_empty():
        raise click.exceptions.Exit(1)
