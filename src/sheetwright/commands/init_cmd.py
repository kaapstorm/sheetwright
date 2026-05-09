"""Implementation of `sheetwright init`."""

from __future__ import annotations

from pathlib import Path

import click

DEFAULT_SHEETWRIGHT_TOML = """\
[project]
name = "my-model"

[build]
calc_engine = "libreoffice"
"""

DEFAULT_WORKBOOK_TOML = """\
[workbook]
name = "my-model"

# Sheets are listed in workbook order. Each entry must match a file in sheets/
# (without the .md/.yaml extension).
sheets = []

# Workbook-scoped named ranges:
# [[named_ranges]]
# name = "growth_rate"
# scope = "workbook"
# ref = "Assumptions!B5"
"""

DEFAULT_GITIGNORE = """\
build/
.sheetwright/
"""


def run(path: str) -> None:
    project = Path(path).resolve()
    if not _project_is_empty(project):
        raise click.ClickException(f'{project} is not empty.')

    project.mkdir(parents=True, exist_ok=True)
    (project / 'sheets').mkdir()
    (project / 'data').mkdir()
    (project / 'tests').mkdir()
    (project / 'tests' / '__init__.py').write_text('')
    (project / 'sheetwright.toml').write_text(DEFAULT_SHEETWRIGHT_TOML)
    (project / 'workbook.toml').write_text(DEFAULT_WORKBOOK_TOML)
    (project / '.gitignore').write_text(DEFAULT_GITIGNORE)

    click.echo(f'Initialised sheetwright project at {project}')


def _project_is_empty(project):
    expected_files = ('.git', '.venv', 'pyproject.toml', 'uv.lock')
    return (
        not project.exists()
        or any(f for f in project.iterdir() if f not in expected_files)
    )
