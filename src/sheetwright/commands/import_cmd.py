"""Implementation of `sheetwright import`."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import click

from sheetwright.exceptions import ProjectError
from sheetwright.project import Project
from sheetwright.reimport import archive_xlsx
from sheetwright.security import SecurityLimits, get_operator_limits
from sheetwright.source.writer import write_source
from sheetwright.xlsx.flatten import (
    detect_external_refs,
    flatten_external_refs,
)
from sheetwright.xlsx.reader import read_xlsx


def run(
    *,
    xlsx_path: Optional[str],
    project_path: str,
    archive: bool,
    flatten: bool,
    non_interactive: bool,
    apply: bool,
    abort: bool,
    force: bool,
) -> None:
    project_root = Path(project_path).resolve()

    try:
        project = Project.open(project_root)
    except ProjectError as e:
        raise click.ClickException(str(e))

    if abort:
        from sheetwright.reimport import clear_session

        clear_session(project.reimport_session_path)
        click.echo('Session cleared.')
        return

    if apply:
        from sheetwright.reimport import apply_session

        apply_session(project, archive=archive, flatten=flatten)
        return

    if xlsx_path is None:
        raise click.ClickException(
            'Missing XLSX argument. Pass a path, or use --apply / --abort '
            'to act on a staged session.'
        )
    xlsx = Path(xlsx_path).resolve()

    if project.has_source():
        from sheetwright.reimport import do_reimport

        do_reimport(
            project,
            xlsx,
            archive=archive,
            flatten=flatten,
            non_interactive=non_interactive,
            force=force,
        )
        return

    limits = SecurityLimits.effective(
        get_operator_limits(), project.config.security
    )
    extrefs = detect_external_refs(xlsx, limits=limits)
    if extrefs and not flatten:
        raise click.ClickException(
            'Workbook contains external references; '
            'pass --flatten to replace them with cached values, '
            'or resolve them in Excel before importing.\n'
            'First few: ' + ', '.join(extrefs[:3])
        )

    wb = read_xlsx(xlsx, limits=limits)
    if flatten:
        flatten_external_refs(wb, xlsx, limits=limits)
    write_source(wb, project_root)

    if archive:
        archive_xlsx(xlsx, project_root)

    click.echo(f'Imported {xlsx} into {project_root}')
