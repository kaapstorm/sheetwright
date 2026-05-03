"""Re-import flow: detect, diff, prompt, apply.

Public entry points:
    do_reimport(project, xlsx, *, archive, flatten, non_interactive, force)
        - full interactive (or session-staging) flow.
    archive_xlsx(xlsx, project_root)
        - copy the imported xlsx into imports/ with a timestamped name.

Tasks 10 and 11 add `apply_session` and the uncommitted-source guard.
"""

from __future__ import annotations

import shutil
import subprocess
from datetime import datetime, timezone
from pathlib import Path

import click

from claudesheets.calc.cache import hash_xlsx
from claudesheets.diff import diff_workbooks
from claudesheets.diff.format import render
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


def has_uncommitted_changes(project_root: Path) -> bool:
    """Return True if `sheets/` has uncommitted changes in git.

    Returns False when the directory is not a git repo (we can't tell
    what's "uncommitted") or when git is unavailable.
    """
    if not (project_root / '.git').is_dir():
        return False
    proc = subprocess.run(
        ['git', 'status', '--porcelain', 'sheets/'],
        cwd=project_root,
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        return False
    return bool(proc.stdout.strip())


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
    report = render(diff)
    click.echo(report, nl=False)

    if diff.is_empty():
        click.echo('Nothing to merge.')
        return

    if non_interactive:
        save_session(
            project.reimport_session_path,
            ReimportSession(
                xlsx_path=str(xlsx),
                xlsx_sha256=hash_xlsx(xlsx),
                diff_summary=report,
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

    write_source(new_wb, project.root)
    if archive:
        archive_xlsx(xlsx, project.root)
    click.echo(f'Source updated from {xlsx}.')


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
