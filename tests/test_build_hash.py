import json
import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from sheetwright.build_hash import (
    BuildHashRecord,
    read_build_hash,
    write_build_hash,
)


@contextmanager
def _tmp_path():
    with tempfile.TemporaryDirectory() as td:
        yield Path(td)


@test
def write_then_read_round_trips():
    with _tmp_path() as tmp_path:
        path = tmp_path / 'bh.json'
        rec = BuildHashRecord(
            name='in', sha256='abc123', built_at='2026-01-01T00:00:00Z'
        )
        write_build_hash(path, rec)
        assert read_build_hash(path) == rec


@test
def read_missing_returns_none():
    with _tmp_path() as tmp_path:
        assert read_build_hash(tmp_path / 'missing.json') is None


@test
def read_corrupt_returns_none():
    with _tmp_path() as tmp_path:
        p = tmp_path / 'corrupt.json'
        p.write_text('{not valid json')
        assert read_build_hash(p) is None


@test
def read_missing_keys_returns_none():
    with _tmp_path() as tmp_path:
        p = tmp_path / 'partial.json'
        p.write_text('{"name": "in"}')  # missing sha256, built_at
        assert read_build_hash(p) is None


@test
def write_creates_parent_dir():
    with _tmp_path() as tmp_path:
        path = tmp_path / 'nested' / 'dir' / 'bh.json'
        rec = BuildHashRecord(name='x', sha256='d', built_at='t')
        write_build_hash(path, rec)
        assert path.is_file()
        data = json.loads(path.read_text())
        assert data['name'] == 'x'
