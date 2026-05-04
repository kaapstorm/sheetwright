"""Implementation of `claudesheets test` (testsweet, in-process).

Uses testsweet's lower-level `discover` + `run` API directly so we
can run user tests without mutating the process's cwd. This matters
for long-running hosts (the MCP server in particular).

User projects lose support for `[tool.testsweet.discovery]` in their
`pyproject.toml` via this command — we walk `tests/test_*.py` (and
honour `targets` if given) and import each file ourselves. Users who
want `[tool.testsweet.discovery]` can `python -m testsweet` from
their project root directly.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from typing import Sequence

import click
from testsweet import (
    Errored,
    Failed,
    Passed,
    Skipped,
    XFailed,
    XPassed,
    run as ts_run,
)

from claudesheets.exceptions import ProjectError
from claudesheets.project import Project


def run(*, project_path: str, targets: Sequence[str]) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    if not project.tests_dir.is_dir():
        raise click.ClickException(f'no tests/ directory in {project.root}')

    test_files = list(_resolve_targets(project, list(targets)))
    if not test_files:
        click.echo('no tests collected')
        return

    saved_path = list(sys.path)
    saved_modules = {
        name: mod
        for name, mod in sys.modules.items()
        if name == '_user_tests' or name.startswith('_user_tests.')
    }
    sys.path.insert(0, str(project.root))
    try:
        any_failure = False
        for test_file in test_files:
            module = _import_module(test_file)
            for name, outcome in ts_run(module):
                full = f'{test_file.relative_to(project.root)}::{name}'
                match outcome:
                    case Passed():
                        click.echo(f'{full} ... ok')
                    case Skipped(reason=reason):
                        click.echo(f'{full} ... skipped: {reason}')
                    case XFailed(reason=reason):
                        click.echo(f'{full} ... xfail: {reason}')
                    case XPassed(reason=reason):
                        any_failure = True
                        click.echo(f'{full} ... XPASS: {reason}')
                    case Failed(exc=exc):
                        any_failure = True
                        click.echo(
                            f'{full} ... FAIL: {type(exc).__name__}: {exc}'
                        )
                    case Errored(exc=exc):
                        any_failure = True
                        click.echo(
                            f'{full} ... ERROR: {type(exc).__name__}: {exc}'
                        )
    finally:
        sys.path[:] = saved_path
        # Restore any pre-existing _user_tests modules we may have
        # shadowed; remove ones we created.
        for name in [
            n
            for n in sys.modules
            if n == '_user_tests' or n.startswith('_user_tests.')
        ]:
            del sys.modules[name]
        sys.modules.update(saved_modules)

    if any_failure:
        raise click.exceptions.Exit(1)


def _resolve_targets(project: Project, targets: list[str]) -> list[Path]:
    """Return the list of test files to run."""
    if not targets:
        return sorted(project.tests_dir.rglob('test_*.py'))
    out = []
    for t in targets:
        p = (project.root / t).resolve()
        if p.is_file():
            out.append(p)
        elif p.is_dir():
            out.extend(sorted(p.rglob('test_*.py')))
        else:
            raise click.ClickException(f'no such target: {t}')
    return out


def _import_module(path: Path):
    """Import a test file under a stable namespace."""
    name = '_user_tests.' + path.stem
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise click.ClickException(f'cannot import {path}')
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module
