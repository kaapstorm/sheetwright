"""Implementation of `claudesheets check`."""

from __future__ import annotations

import click

from claudesheets.diff.check import check_workbook
from claudesheets.exceptions import ProjectError
from claudesheets.project import Project
from claudesheets.source.reader import read_source


def run(*, project_path: str) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    wb = read_source(project.root)
    issues = check_workbook(wb, project)

    if not issues:
        click.echo('no issues')
        return

    for issue in issues:
        loc = f' at {issue.location}' if issue.location else ''
        click.echo(f'  [{issue.kind}]{loc}: {issue.detail}')
    raise click.exceptions.Exit(1)
