"""Unit tests for SecurityLimits and OperatorLimits."""

import os
from unittest.mock import patch

from testsweet import test

from sheetwright.security import (
    OperatorLimits,
    SecurityLimits,
    _reset_operator_limits,
    get_operator_limits,
)


def _operator(**kwargs: object) -> OperatorLimits:
    defaults = dict(
        max_xlsx_uncompressed_bytes=200 * 1024 * 1024,
        max_xlsx_sheet_count=200,
        max_xlsx_cells_per_sheet=5_000_000,
        max_xlsx_shared_strings=5_000_000,
        soffice_timeout=120.0,
    )
    defaults.update(kwargs)  # type: ignore[arg-type]
    return OperatorLimits(**defaults)  # type: ignore[arg-type]


def _project(**kwargs: object) -> SecurityLimits:
    defaults = dict(
        max_xlsx_uncompressed_bytes=200 * 1024 * 1024,
        max_xlsx_sheet_count=200,
        max_xlsx_cells_per_sheet=5_000_000,
        max_xlsx_shared_strings=5_000_000,
        soffice_timeout=120.0,
    )
    defaults.update(kwargs)  # type: ignore[arg-type]
    return SecurityLimits(**defaults)  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# effective() — operator value picked when stricter (one assertion per field)
# ---------------------------------------------------------------------------


@test
def effective_picks_operator_bytes_when_stricter():
    op = _operator(max_xlsx_uncompressed_bytes=50)
    proj = _project(max_xlsx_uncompressed_bytes=100)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_uncompressed_bytes == 50


@test
def effective_picks_operator_sheet_count_when_stricter():
    op = _operator(max_xlsx_sheet_count=10)
    proj = _project(max_xlsx_sheet_count=50)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_sheet_count == 10


@test
def effective_picks_operator_cells_per_sheet_when_stricter():
    op = _operator(max_xlsx_cells_per_sheet=1_000)
    proj = _project(max_xlsx_cells_per_sheet=5_000)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_cells_per_sheet == 1_000


@test
def effective_picks_operator_shared_strings_when_stricter():
    op = _operator(max_xlsx_shared_strings=999)
    proj = _project(max_xlsx_shared_strings=9_999)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_shared_strings == 999


@test
def effective_picks_operator_soffice_timeout_when_stricter():
    op = _operator(soffice_timeout=30.0)
    proj = _project(soffice_timeout=90.0)
    result = SecurityLimits.effective(op, proj)
    assert result.soffice_timeout == 30.0


# ---------------------------------------------------------------------------
# effective() — project value picked when stricter
# ---------------------------------------------------------------------------


@test
def effective_picks_project_bytes_when_stricter():
    op = _operator(max_xlsx_uncompressed_bytes=100)
    proj = _project(max_xlsx_uncompressed_bytes=50)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_uncompressed_bytes == 50


@test
def effective_picks_project_sheet_count_when_stricter():
    op = _operator(max_xlsx_sheet_count=50)
    proj = _project(max_xlsx_sheet_count=10)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_sheet_count == 10


@test
def effective_picks_project_cells_per_sheet_when_stricter():
    op = _operator(max_xlsx_cells_per_sheet=5_000)
    proj = _project(max_xlsx_cells_per_sheet=1_000)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_cells_per_sheet == 1_000


@test
def effective_picks_project_shared_strings_when_stricter():
    op = _operator(max_xlsx_shared_strings=9_999)
    proj = _project(max_xlsx_shared_strings=999)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_shared_strings == 999


@test
def effective_picks_project_soffice_timeout_when_stricter():
    op = _operator(soffice_timeout=90.0)
    proj = _project(soffice_timeout=30.0)
    result = SecurityLimits.effective(op, proj)
    assert result.soffice_timeout == 30.0


# ---------------------------------------------------------------------------
# effective() — project is None → returns operator values
# ---------------------------------------------------------------------------


@test
def effective_returns_operator_when_project_is_none():
    op = _operator(
        max_xlsx_uncompressed_bytes=42,
        max_xlsx_sheet_count=7,
        max_xlsx_cells_per_sheet=8,
        max_xlsx_shared_strings=9,
        soffice_timeout=11.0,
    )
    result = SecurityLimits.effective(op, None)
    assert result.max_xlsx_uncompressed_bytes == 42
    assert result.max_xlsx_sheet_count == 7
    assert result.max_xlsx_cells_per_sheet == 8
    assert result.max_xlsx_shared_strings == 9
    assert result.soffice_timeout == 11.0


# ---------------------------------------------------------------------------
# effective() — equal values → returns equal
# ---------------------------------------------------------------------------


@test
def effective_returns_equal_when_values_equal():
    op = _operator(max_xlsx_uncompressed_bytes=100)
    proj = _project(max_xlsx_uncompressed_bytes=100)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_uncompressed_bytes == 100


# ---------------------------------------------------------------------------
# effective() — per-field independence
# ---------------------------------------------------------------------------


@test
def effective_per_field_independent():
    # operator tighter on sheets, project tighter on cells
    op = _operator(max_xlsx_sheet_count=5, max_xlsx_cells_per_sheet=10_000)
    proj = _project(max_xlsx_sheet_count=50, max_xlsx_cells_per_sheet=1_000)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_sheet_count == 5
    assert result.max_xlsx_cells_per_sheet == 1_000


# ---------------------------------------------------------------------------
# effective() — mixed: project tighter on bytes, operator tighter on sheets
# ---------------------------------------------------------------------------


@test
def effective_mixed_tighter_per_field():
    op = _operator(max_xlsx_uncompressed_bytes=200, max_xlsx_sheet_count=10)
    proj = _project(max_xlsx_uncompressed_bytes=50, max_xlsx_sheet_count=100)
    result = SecurityLimits.effective(op, proj)
    assert result.max_xlsx_uncompressed_bytes == 50
    assert result.max_xlsx_sheet_count == 10


# ---------------------------------------------------------------------------
# OperatorLimits.from_environment() — reads each env var
# ---------------------------------------------------------------------------

_CLEAR_ENV = {
    'SHEETWRIGHT_MAX_XLSX_BYTES': None,
    'SHEETWRIGHT_MAX_XLSX_SHEETS': None,
    'SHEETWRIGHT_MAX_XLSX_CELLS_PER_SHEET': None,
    'SHEETWRIGHT_MAX_XLSX_SHARED_STRINGS': None,
    'SHEETWRIGHT_SOFFICE_TIMEOUT': None,
}


@test
def from_environment_reads_each_var():
    env = {
        'SHEETWRIGHT_MAX_XLSX_BYTES': '12345',
        'SHEETWRIGHT_MAX_XLSX_SHEETS': '7',
        'SHEETWRIGHT_MAX_XLSX_CELLS_PER_SHEET': '99',
        'SHEETWRIGHT_MAX_XLSX_SHARED_STRINGS': '88',
        'SHEETWRIGHT_SOFFICE_TIMEOUT': '45.5',
    }
    with patch.dict(os.environ, env):
        limits = OperatorLimits.from_environment()
    assert limits.max_xlsx_uncompressed_bytes == 12345
    assert limits.max_xlsx_sheet_count == 7
    assert limits.max_xlsx_cells_per_sheet == 99
    assert limits.max_xlsx_shared_strings == 88
    assert limits.soffice_timeout == 45.5


@test
def from_environment_falls_back_on_missing():
    # Remove all SHEETWRIGHT_* vars to test pure defaults
    clean = {k: v for k, v in os.environ.items() if 'SHEETWRIGHT' not in k}
    with patch.dict(os.environ, clean, clear=True):
        limits = OperatorLimits.from_environment()
    defaults = SecurityLimits.defaults()
    assert (
        limits.max_xlsx_uncompressed_bytes
        == defaults.max_xlsx_uncompressed_bytes
    )
    assert limits.max_xlsx_sheet_count == defaults.max_xlsx_sheet_count
    assert limits.max_xlsx_cells_per_sheet == defaults.max_xlsx_cells_per_sheet
    assert limits.max_xlsx_shared_strings == defaults.max_xlsx_shared_strings
    assert limits.soffice_timeout == defaults.soffice_timeout


@test
def from_environment_falls_back_per_field_on_unparseable():
    env = {
        'SHEETWRIGHT_MAX_XLSX_BYTES': 'abc',
        'SHEETWRIGHT_MAX_XLSX_SHEETS': 'not-a-number',
        'SHEETWRIGHT_MAX_XLSX_CELLS_PER_SHEET': '',
        'SHEETWRIGHT_MAX_XLSX_SHARED_STRINGS': 'not-an-int',
        'SHEETWRIGHT_SOFFICE_TIMEOUT': 'bad',
    }
    with patch.dict(os.environ, env):
        limits = OperatorLimits.from_environment()
    defaults = SecurityLimits.defaults()
    assert (
        limits.max_xlsx_uncompressed_bytes
        == defaults.max_xlsx_uncompressed_bytes
    )
    assert limits.max_xlsx_sheet_count == defaults.max_xlsx_sheet_count
    assert limits.max_xlsx_cells_per_sheet == defaults.max_xlsx_cells_per_sheet
    assert limits.max_xlsx_shared_strings == defaults.max_xlsx_shared_strings
    assert limits.soffice_timeout == defaults.soffice_timeout


# ---------------------------------------------------------------------------
# get_operator_limits() — lazy init from environment
# ---------------------------------------------------------------------------


@test
def get_operator_limits_returns_value_from_env():
    _reset_operator_limits()
    try:
        env = {'SHEETWRIGHT_MAX_XLSX_SHEETS': '42'}
        with patch.dict(os.environ, env):
            limits = get_operator_limits()
        assert limits.max_xlsx_sheet_count == 42
    finally:
        _reset_operator_limits()
