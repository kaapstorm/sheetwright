"""Small git-related helpers."""

from __future__ import annotations

import subprocess
from pathlib import Path


def has_uncommitted_changes(
    project_root: Path, subpath: str = 'sheets/'
) -> bool:
    """Return True if `subpath` (relative to project_root) has
    uncommitted changes in git.

    Returns False when the directory is not a git repo (we can't
    tell what's "uncommitted") or when git is unavailable.
    """
    if not (project_root / '.git').is_dir():
        return False
    proc = subprocess.run(
        ['git', 'status', '--porcelain', subpath],
        cwd=project_root,
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        return False
    return bool(proc.stdout.strip())
