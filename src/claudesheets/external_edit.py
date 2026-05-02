"""Detect whether the most-recent built xlsx has been edited
externally (i.e., between claudesheets sessions).
"""

from __future__ import annotations

import click

from claudesheets.build_hash import read_build_hash
from claudesheets.calc.cache import hash_xlsx
from claudesheets.config import load_project
from claudesheets.project import Project


def warn_if_externally_edited(project: Project) -> None:
    """Emit a Click warning when the built xlsx no longer matches the
    recorded build hash. No-op when there's no record yet, or no
    built xlsx, or hashes match.
    """
    rec = read_build_hash(project.build_hash_path)
    if rec is None:
        return
    cfg = load_project(project.claudesheets_toml.read_text())
    built = project.build_dir / f'{cfg.name}.xlsx'
    if not built.is_file():
        return
    current = hash_xlsx(built)
    if current == rec.sha256:
        return
    click.echo(
        f'WARNING: {built} has been modified externally. '
        f'Run `claudesheets import {built}` to review changes.',
        err=True,
    )
