from pathlib import Path

from claudesheets.reimport.session import (
    ReimportSession,
    clear_session,
    load_session,
    save_session,
)


def test_session_round_trips(tmp_path: Path):
    p = tmp_path / 'session.json'
    s = ReimportSession(
        xlsx_path='/tmp/in.xlsx',
        xlsx_sha256='abc123',
        diff_summary='2 cells changed',
        created_at='2026-05-02T12:00:00Z',
    )
    save_session(p, s)
    assert load_session(p) == s


def test_load_missing_returns_none(tmp_path: Path):
    assert load_session(tmp_path / 'absent.json') is None


def test_clear_session_removes_file(tmp_path: Path):
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


def test_clear_session_no_file_is_noop(tmp_path: Path):
    clear_session(tmp_path / 'never.json')  # must not raise
