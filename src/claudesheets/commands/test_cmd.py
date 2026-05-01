"""Implementation of `claudesheets test` (testsweet, in-process).

We invoke testsweet's `main(argv) -> int` programmatically. testsweet
saves and restores `sys.path` itself; we additionally save/restore
`os.getcwd()` because testsweet reads `[tool.testsweet.discovery]`
config from the cwd's `pyproject.toml`.
"""

from __future__ import annotations

import os
import sys
from typing import Sequence

import click

from claudesheets.exceptions import ProjectError
from claudesheets.project import Project


def run(*, project_path: str, targets: Sequence[str]) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    if not project.tests_dir.is_dir():
        raise click.ClickException(f'no tests/ directory in {project.root}')

    from testsweet.__main__ import main as testsweet_main

    argv = list(targets) if targets else [str(project.tests_dir)]
    prev_cwd = os.getcwd()
    # The host process (e.g. pytest running claudesheets' own tests)
    # may have a `tests` package cached under that name. The user
    # project's tests/ would resolve to the same dotted name and
    # collide. Drop any pre-existing `tests`/`tests.*` modules so
    # testsweet's importlib lookup finds the project's tree, and
    # restore them afterwards.
    saved_modules = {
        name: mod
        for name, mod in sys.modules.items()
        if name == 'tests' or name.startswith('tests.')
    }
    for name in saved_modules:
        del sys.modules[name]
    try:
        os.chdir(project.root)
        rc = testsweet_main(argv)
    finally:
        os.chdir(prev_cwd)
        for name in [
            n for n in sys.modules if n == 'tests' or n.startswith('tests.')
        ]:
            del sys.modules[name]
        sys.modules.update(saved_modules)

    if rc != 0:
        raise click.exceptions.Exit(rc)
