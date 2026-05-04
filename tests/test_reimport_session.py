import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from claudesheets.reimport.session import (
    ReimportSession,
    clear_session,
    load_session,
    save_session,
)


@contextmanager
def _tmp_path():
    with tempfile.TemporaryDirectory() as td:
        yield Path(td)


@test
def session_round_trips():
    with _tmp_path() as tmp_path:
        p = tmp_path / 'session.json'
        s = ReimportSession(
            xlsx_path='/tmp/in.xlsx',
            xlsx_sha256='abc123',
            diff_summary='2 cells changed',
            created_at='2026-05-02T12:00:00Z',
        )
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
        save_session(
            p,
            ReimportSession(
                xlsx_path='/x',
                xlsx_sha256='a',
                diff_summary='',
                created_at='t',
            ),
        )
        clear_session(p)
        assert not p.exists()


@test
def clear_session_no_file_is_noop():
    with _tmp_path() as tmp_path:
        clear_session(tmp_path / 'never.json')  # must not raise
