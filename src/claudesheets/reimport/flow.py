"""Re-import flow: detect, diff, prompt, apply.

Public entry points:
    stage_reimport(project, xlsx, *, flatten, force) -> StagedReimport
        - Click-free: validates, computes diff, returns staged result.
    commit_staged(project, staged, *, archive)
        - Click-free: writes source from a StagedReimport.
    do_reimport(project, xlsx, *, archive, flatten, non_interactive, force)
        - thin Click-aware wrapper around the above (prompts + echoes).
    apply_session(project, *, archive, flatten)
        - complete a previously-staged -I session.
    archive_xlsx(xlsx, project_root)
        - copy the imported xlsx into imports/ with a timestamped name.
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

import click

from claudesheets.calc.cache import hash_xlsx
from claudesheets.diff import diff_workbooks
from claudesheets.diff.format import render
from claudesheets.diff.model import WorkbookDiff
from claudesheets.gitutil import has_uncommitted_changes
from claudesheets.model.workbook import Workbook
from claudesheets.project import Project
from claudesheets.reimport.session import (
    ReimportSession,
    clear_session,
    load_session,
    save_session,
)
from claudesheets.source.reader import read_source
from claudesheets.source.writer import write_source
from claudesheets.xlsx.flatten import (
    detect_external_refs,
    flatten_external_refs,
)
from claudesheets.xlsx.reader import read_xlsx


@dataclass(frozen=True)
class StagedReimport:
    diff: WorkbookDiff
    rendered_diff: str
    new_workbook: Workbook
    xlsx_path: Path
    xlsx_sha256: str


def stage_reimport(
    project: Project,
    xlsx: Path,
    *,
    flatten: bool,
    force: bool,
) -> StagedReimport:
    """Click-free re-import staging.

    Validates the uncommitted-source guard, optionally flattens
    external refs, computes the diff between the current source and
    the new xlsx, and returns it. No I/O beyond reading the xlsx.

    Raises `click.ClickException` for user-facing errors (uncommitted
    changes, external refs without flatten). Callers translate as
    appropriate (CLI: print and exit; MCP: surface as typed error).
    """
    if has_uncommitted_changes(project.root) and not force:
        raise click.ClickException(
            'You have uncommitted changes in sheets/. Commit or stash '
            'them before importing, or pass --force to discard.'
        )

    extrefs = detect_external_refs(xlsx)
    if extrefs and not flatten:
        raise click.ClickException(
            'Workbook contains external references; '
            'pass --flatten to replace them with cached values.'
        )

    new_wb = read_xlsx(xlsx)
    if flatten:
        flatten_external_refs(new_wb, xlsx)

    current_wb = read_source(project.root)
    diff = diff_workbooks(current_wb, new_wb)
    rendered = render(diff)

    return StagedReimport(
        diff=diff,
        rendered_diff=rendered,
        new_workbook=new_wb,
        xlsx_path=Path(xlsx).resolve(),
        xlsx_sha256=hash_xlsx(xlsx),
    )


def commit_staged(
    project: Project, staged: StagedReimport, *, archive: bool
) -> None:
    """Click-free apply: write source from a StagedReimport."""
    write_source(staged.new_workbook, project.root)
    if archive:
        archive_xlsx(staged.xlsx_path, project.root)


def do_reimport(
    project: Project,
    xlsx: Path,
    *,
    archive: bool,
    flatten: bool,
    non_interactive: bool,
    force: bool,
) -> None:
    """Run the re-import flow against an already-populated source.

    `force=True` skips the uncommitted-source guard ONLY. It does NOT
    bypass the external-reference check (use `--flatten` for that),
    nor does it auto-overwrite without prompting; the user still
    chooses Merge / Overwrite / Reject (or stages with -I).
    """
    staged = stage_reimport(project, xlsx, flatten=flatten, force=force)
    click.echo(staged.rendered_diff)

    if staged.diff.is_empty():
        click.echo('Nothing to merge.')
        return

    if non_interactive:
        save_session(
            project.reimport_session_path,
            ReimportSession(
                xlsx_path=str(staged.xlsx_path),
                xlsx_sha256=staged.xlsx_sha256,
                diff_summary=staged.rendered_diff,
                created_at=datetime.now(timezone.utc).isoformat(),
            ),
        )
        click.echo(
            'Run `claudesheets import --apply` to apply, '
            'or `--abort` to discard.'
        )
        return

    choice = click.prompt(
        'Apply changes? [m]erge / [o]verwrite / [r]eject',
        type=click.Choice(['m', 'o', 'r'], case_sensitive=False),
        default='r',
    ).lower()

    if choice == 'r':
        click.echo('Rejected; source unchanged.')
        return

    commit_staged(project, staged, archive=archive)
    click.echo(f'Source updated from {staged.xlsx_path}.')


def apply_session(project: Project, *, archive: bool, flatten: bool) -> None:
    """Complete a previously-staged -I session.

    Reads the session, opens the staged xlsx, optionally flattens
    external refs, writes source, optionally archives, and clears
    the session file.
    """
    session = load_session(project.reimport_session_path)
    if session is None:
        raise click.ClickException(
            'No staged re-import session. '
            'Run `claudesheets import <xlsx> -I` first.'
        )

    xlsx = Path(session.xlsx_path)
    if not xlsx.is_file():
        raise click.ClickException(
            f'Staged xlsx no longer exists at {xlsx}. Re-stage with -I.'
        )

    current_hash = hash_xlsx(xlsx)
    if current_hash != session.xlsx_sha256:
        raise click.ClickException(
            f'Staged xlsx at {xlsx} has been modified since `-I` ('
            f'recorded hash {session.xlsx_sha256[:12]}, current '
            f'{current_hash[:12]}). Re-stage with `-I` to see the new '
            f'diff.'
        )

    new_wb = read_xlsx(xlsx)
    if flatten:
        flatten_external_refs(new_wb, xlsx)
    write_source(new_wb, project.root)
    if archive:
        archive_xlsx(xlsx, project.root)

    clear_session(project.reimport_session_path)
    click.echo(f'Applied staged changes from {xlsx}.')


def archive_xlsx(xlsx: Path, project_root: Path) -> None:
    imports_dir = project_root / 'imports'
    imports_dir.mkdir(exist_ok=True)
    ts = datetime.now().strftime('%Y-%m-%dT%H%M')
    shutil.copy2(xlsx, imports_dir / f'{ts}.xlsx')
