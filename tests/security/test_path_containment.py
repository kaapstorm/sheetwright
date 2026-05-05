"""Unit tests for resolve_under: containment, symlinks, non-existent paths."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

from testsweet import catch_exceptions, test

from sheetwright.security import PathOutsideProjectError, resolve_under


def _can_create_symlinks() -> bool:
    try:
        with tempfile.TemporaryDirectory() as td:
            src = Path(td) / 'target.txt'
            src.write_text('x')
            link = Path(td) / 'link.txt'
            link.symlink_to(src)
        return True
    except (OSError, NotImplementedError):
        return False


@test
def relative_path_under_root_is_accepted():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        sub = root / 'sub'
        sub.mkdir()
        result = resolve_under(root, 'sub/file.xlsx')
        assert result == (root / 'sub' / 'file.xlsx').resolve()


@test
def relative_dotdot_escape_raises():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        with catch_exceptions() as excs:
            resolve_under(root, '../escape.xlsx')
        assert excs and isinstance(excs[0], PathOutsideProjectError)


@test
def absolute_path_outside_root_raises():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        with catch_exceptions() as excs:
            resolve_under(root, '/etc/passwd')
        assert excs and isinstance(excs[0], PathOutsideProjectError)


@test
def symlink_under_root_pointing_outside_raises():
    if not _can_create_symlinks():
        return  # Skip: platform does not support symlinks
    with (
        tempfile.TemporaryDirectory() as td1,
        tempfile.TemporaryDirectory() as td2,
    ):
        root = Path(td1)
        outside_file = Path(td2) / 'secret.txt'
        outside_file.write_text('secret')
        link = root / 'evil_link.txt'
        link.symlink_to(outside_file)
        with catch_exceptions() as excs:
            resolve_under(root, 'evil_link.txt')
        assert excs and isinstance(excs[0], PathOutsideProjectError)


@test
def symlink_under_root_pointing_inside_is_accepted():
    if not _can_create_symlinks():
        return  # Skip: platform does not support symlinks
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        real_file = root / 'real.txt'
        real_file.write_text('ok')
        link = root / 'link_to_real.txt'
        link.symlink_to(real_file)
        result = resolve_under(root, 'link_to_real.txt')
        assert result == real_file.resolve()


@test
def non_existent_path_under_root_is_accepted():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        # 'newdir' does not exist; the function must handle this gracefully.
        result = resolve_under(root, 'newdir/newfile.xlsx')
        assert result == root / 'newdir' / 'newfile.xlsx'
        assert not result.exists()
