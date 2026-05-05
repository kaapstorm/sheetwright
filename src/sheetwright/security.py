"""Security limits for sheetwright operations."""

from __future__ import annotations

import os
from dataclasses import dataclass


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
        if project is None:
            return cls(
                max_xlsx_uncompressed_bytes=operator.max_xlsx_uncompressed_bytes,
                max_xlsx_sheet_count=operator.max_xlsx_sheet_count,
                max_xlsx_cells_per_sheet=operator.max_xlsx_cells_per_sheet,
                max_xlsx_shared_strings=operator.max_xlsx_shared_strings,
                soffice_timeout=operator.soffice_timeout,
            )
        return cls(
            max_xlsx_uncompressed_bytes=min(
                operator.max_xlsx_uncompressed_bytes,
                project.max_xlsx_uncompressed_bytes,
            ),
            max_xlsx_sheet_count=min(
                operator.max_xlsx_sheet_count,
                project.max_xlsx_sheet_count,
            ),
            max_xlsx_cells_per_sheet=min(
                operator.max_xlsx_cells_per_sheet,
                project.max_xlsx_cells_per_sheet,
            ),
            max_xlsx_shared_strings=min(
                operator.max_xlsx_shared_strings,
                project.max_xlsx_shared_strings,
            ),
            soffice_timeout=min(
                operator.soffice_timeout,
                project.soffice_timeout,
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
