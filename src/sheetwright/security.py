"""Security limits for sheetwright operations."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


class PathOutsideProjectError(Exception):
    """Raised when a candidate path resolves outside the project root."""


def resolve_under(root: Path, candidate: str | Path) -> Path:
    """Resolve `candidate` relative to `root`; verify the resolved
    path is under `root` after symlink collapse. Returns the resolved
    Path. Raises PathOutsideProjectError if it escapes.

    Handles non-existent paths (e.g. an output file that hasn't been
    written yet) by resolving the deepest existing ancestor and
    re-attaching the missing suffix.
    """
    root_resolved = root.resolve(strict=True)

    if Path(candidate).is_absolute():
        target = Path(candidate)
    else:
        target = root_resolved / candidate

    # Walk parents until we find an existing ancestor; resolve that;
    # re-attach the missing tail to handle non-existent paths.
    parts: list[str] = []
    current = target
    while True:
        if current.exists():
            resolved = current.resolve()
            # Re-attach the non-existent suffix
            for part in reversed(parts):
                resolved = resolved / part
            break
        parts.append(current.name)
        parent = current.parent
        if parent == current:
            # Reached filesystem root without finding an existing node;
            # fall back to resolving what we have (no symlink collapsing).
            resolved = target
            break
        current = parent

    if not resolved.is_relative_to(root_resolved):
        raise PathOutsideProjectError(
            f'{candidate!r} resolves outside {root_resolved!r}'
        )
    return resolved


_DEFAULT_MAX_XLSX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024  # 200 MiB
_DEFAULT_MAX_XLSX_SHEET_COUNT = 200
_DEFAULT_MAX_XLSX_CELLS_PER_SHEET = 5_000_000
_DEFAULT_MAX_XLSX_SHARED_STRINGS = 5_000_000
_DEFAULT_SOFFICE_TIMEOUT = 120.0


@dataclass(frozen=True)
class SecurityLimits:
    """Per-project security limits. All fields are maximums (or timeouts).

    Constructed from the ``[security]`` block in ``sheetwright.toml``.
    Project limits may only tighten (lower) the operator ceiling, never
    raise it — enforced by ``SecurityLimits.effective``.
    """

    max_xlsx_uncompressed_bytes: int
    max_xlsx_sheet_count: int
    max_xlsx_cells_per_sheet: int
    max_xlsx_shared_strings: int
    soffice_timeout: float

    @classmethod
    def defaults(cls) -> 'SecurityLimits':
        return cls(
            max_xlsx_uncompressed_bytes=_DEFAULT_MAX_XLSX_UNCOMPRESSED_BYTES,
            max_xlsx_sheet_count=_DEFAULT_MAX_XLSX_SHEET_COUNT,
            max_xlsx_cells_per_sheet=_DEFAULT_MAX_XLSX_CELLS_PER_SHEET,
            max_xlsx_shared_strings=_DEFAULT_MAX_XLSX_SHARED_STRINGS,
            soffice_timeout=_DEFAULT_SOFFICE_TIMEOUT,
        )

    @classmethod
    def effective(
        cls,
        operator: 'OperatorLimits',
        project: 'SecurityLimits | None',
    ) -> 'SecurityLimits':
        """Per-field min(operator, project).

        Project may only tighten the operator ceiling, never raise it.
        If ``project`` is ``None``, returns operator values unchanged.
        """
        p = project if project is not None else operator
        return cls(
            max_xlsx_uncompressed_bytes=min(
                operator.max_xlsx_uncompressed_bytes,
                p.max_xlsx_uncompressed_bytes,
            ),
            max_xlsx_sheet_count=min(
                operator.max_xlsx_sheet_count,
                p.max_xlsx_sheet_count,
            ),
            max_xlsx_cells_per_sheet=min(
                operator.max_xlsx_cells_per_sheet,
                p.max_xlsx_cells_per_sheet,
            ),
            max_xlsx_shared_strings=min(
                operator.max_xlsx_shared_strings,
                p.max_xlsx_shared_strings,
            ),
            soffice_timeout=min(
                operator.soffice_timeout,
                p.soffice_timeout,
            ),
        )


@dataclass(frozen=True)
class OperatorLimits:
    """Process-wide ceiling sourced from env / mcp flags."""

    max_xlsx_uncompressed_bytes: int
    max_xlsx_sheet_count: int
    max_xlsx_cells_per_sheet: int
    max_xlsx_shared_strings: int
    soffice_timeout: float

    @classmethod
    def from_environment(cls) -> 'OperatorLimits':
        """Read SHEETWRIGHT_MAX_* env vars; fall back to defaults
        per-field on missing or unparseable."""
        return cls(
            max_xlsx_uncompressed_bytes=_read_int_env(
                'SHEETWRIGHT_MAX_XLSX_BYTES',
                _DEFAULT_MAX_XLSX_UNCOMPRESSED_BYTES,
            ),
            max_xlsx_sheet_count=_read_int_env(
                'SHEETWRIGHT_MAX_XLSX_SHEETS',
                _DEFAULT_MAX_XLSX_SHEET_COUNT,
            ),
            max_xlsx_cells_per_sheet=_read_int_env(
                'SHEETWRIGHT_MAX_XLSX_CELLS_PER_SHEET',
                _DEFAULT_MAX_XLSX_CELLS_PER_SHEET,
            ),
            max_xlsx_shared_strings=_read_int_env(
                'SHEETWRIGHT_MAX_XLSX_SHARED_STRINGS',
                _DEFAULT_MAX_XLSX_SHARED_STRINGS,
            ),
            soffice_timeout=_read_float_env(
                'SHEETWRIGHT_SOFFICE_TIMEOUT',
                _DEFAULT_SOFFICE_TIMEOUT,
            ),
        )


_operator_limits: OperatorLimits | None = None


def get_operator_limits() -> OperatorLimits:
    """Process-wide operator limits.

    First call reads env vars and caches; subsequent calls return
    the cached value. Idempotent because `from_environment()` is
    pure with respect to its inputs.
    """
    global _operator_limits
    if _operator_limits is None:
        _operator_limits = OperatorLimits.from_environment()
    return _operator_limits


def _reset_operator_limits() -> None:
    """Reset the cached operator limits. For tests only."""
    global _operator_limits
    _operator_limits = None


def _read_int_env(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except (ValueError, TypeError):
        return default


def _read_float_env(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        return float(raw)
    except (ValueError, TypeError):
        return default
