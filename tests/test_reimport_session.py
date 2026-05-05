import json
import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from sheetwright.exceptions import StaleSessionFormatError
from sheetwright.reimport.session import (
    ReimportSession,
    clear_session,
    load_session,
    save_session,
)


@contextmanager
def _tmp_path():
    with tempfile.TemporaryDirectory() as td:
        yield Path(td)


def _make_session(**overrides) -> ReimportSession:
    defaults = dict(
        xlsx_path='/tmp/in.xlsx',
        xlsx_sha256='abc123',
        diff_summary='2 cells changed',
        created_at='2026-05-02T12:00:00Z',
        staged_filename='abc123.xlsx',
        original_xlsx_path='/tmp/in.xlsx',
    )
    defaults.update(overrides)
    return ReimportSession(**defaults)


@test
def session_round_trips():
    with _tmp_path() as tmp_path:
        p = tmp_path / 'session.json'
        s = _make_session()
        save_session(p, s)
        assert load_session(p) == s


@test
def load_missing_returns_none():
    with _tmp_path() as tmp_path:
        assert load_session(tmp_path / 'absent.json') is None


@test
def clear_session_removes_file():
    with _tmp_path() as tmp_path:
        p = tmp_path / 'session.json'
        save_session(p, _make_session())
        clear_session(p)
        assert not p.exists()


@test
def clear_session_no_file_is_noop():
    with _tmp_path() as tmp_path:
        clear_session(tmp_path / 'never.json')  # must not raise


@test
def load_old_format_session_raises_stale_error():
    """A session JSON lacking staged_filename raises StaleSessionFormatError."""
    with _tmp_path() as tmp_path:
        p = tmp_path / 'session.json'
        old_data = {
            'xlsx_path': '/tmp/in.xlsx',
            'xlsx_sha256': 'abc123',
            'diff_summary': '2 cells changed',
            'created_at': '2026-05-02T12:00:00Z',
        }
        p.write_text(json.dumps(old_data))
        try:
            load_session(p)
            assert False, 'Expected StaleSessionFormatError'
        except StaleSessionFormatError as e:
            assert 're-stage' in str(e).lower()
