"""Implementation of `claudesheets import`."""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import click

from claudesheets.exceptions import ProjectError
from claudesheets.project import Project
from claudesheets.source.writer import write_source
from claudesheets.xlsx.flatten import (
    detect_external_refs,
    flatten_external_refs,
)
from claudesheets.xlsx.reader import read_xlsx


def run(
    *,
    xlsx_path: str,
    project_path: str,
    archive: bool,
    flatten: bool,
    non_interactive: bool,
) -> None:
    xlsx = Path(xlsx_path).resolve()
    project_root = Path(project_path).resolve()

    try:
        project = Project.open(project_root)
    except ProjectError as e:
        raise click.ClickException(str(e))

    sheets_dir = project_root / 'sheets'
    if any(sheets_dir.iterdir()):
        from claudesheets.reimport import do_reimport

        do_reimport(
            project,
            xlsx,
            archive=archive,
            flatten=flatten,
            non_interactive=non_interactive,
        )
        return

    extrefs = detect_external_refs(xlsx)
    if extrefs and not flatten:
        raise click.ClickException(
            'Workbook contains external references; '
            'pass --flatten to replace them with cached values, '
            'or resolve them in Excel before importing.\n'
            'First few: ' + ', '.join(extrefs[:3])
        )

    wb = read_xlsx(xlsx)
    if flatten:
        flatten_external_refs(wb, xlsx)
    write_source(wb, project_root)

    if archive:
        imports_dir = project_root / 'imports'
        imports_dir.mkdir(exist_ok=True)
        ts = datetime.now().strftime('%Y-%m-%dT%H%M')
        shutil.copy2(xlsx, imports_dir / f'{ts}.xlsx')

    click.echo(f'Imported {xlsx} into {project_root}')
