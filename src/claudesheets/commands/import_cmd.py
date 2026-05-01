"""Implementation of `claudesheets import`."""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import click

from claudesheets.exceptions import ProjectError
from claudesheets.project import Project
from claudesheets.source.writer import write_source
from claudesheets.xlsx.reader import read_xlsx


def _detect_external_refs(xlsx_path: Path) -> list[str]:
    import openpyxl

    src = openpyxl.load_workbook(xlsx_path, data_only=False)
    found: set[str] = set()
    for ws in src.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if (
                    isinstance(v, str)
                    and v.startswith('=')
                    and '[' in v
                    and ']' in v
                ):
                    found.add(v)
    return sorted(found)


def run(*, xlsx_path: str, project_path: str, archive: bool) -> None:
    xlsx = Path(xlsx_path).resolve()
    project_root = Path(project_path).resolve()

    try:
        Project.open(project_root)
    except ProjectError as e:
        raise click.ClickException(str(e))

    sheets_dir = project_root / 'sheets'
    if any(sheets_dir.iterdir()):
        raise click.ClickException(
            f'sheets/ in {project_root} is non-empty; '
            'review-first re-import is deferred to a later release.'
        )

    extrefs = _detect_external_refs(xlsx)
    if extrefs:
        raise click.ClickException(
            'Workbook contains external references; '
            'resolve them in Excel before importing.\n'
            'First few: ' + ', '.join(extrefs[:3])
        )

    wb = read_xlsx(xlsx)
    write_source(wb, project_root)

    if archive:
        imports_dir = project_root / 'imports'
        imports_dir.mkdir(exist_ok=True)
        ts = datetime.now().strftime('%Y-%m-%dT%H%M')
        shutil.copy2(xlsx, imports_dir / f'{ts}.xlsx')

    click.echo(f'Imported {xlsx} into {project_root}')
