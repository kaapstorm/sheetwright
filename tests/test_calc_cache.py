import tempfile
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from claudesheets.calc.cache import (
    cache_path_for,
    hash_xlsx,
    read_cached,
    write_cached,
)


@contextmanager
def _tmp_path():
    with tempfile.TemporaryDirectory() as td:
        yield Path(td)


@test
def hash_xlsx_is_stable():
    with _tmp_path() as tmp_path:
        p = tmp_path / 'a.xlsx'
        p.write_bytes(b'hello')
        h1 = hash_xlsx(p)
        h2 = hash_xlsx(p)
        assert h1 == h2
        assert len(h1) == 64  # sha256 hex


@test
def hash_xlsx_changes_with_content():
    with _tmp_path() as tmp_path:
        a = tmp_path / 'a.xlsx'
        b = tmp_path / 'b.xlsx'
        a.write_bytes(b'hello')
        b.write_bytes(b'world')
        assert hash_xlsx(a) != hash_xlsx(b)


@test
def cache_round_trips():
    with _tmp_path() as tmp_path:
        cache_dir = tmp_path / 'cache'
        result = {'Sheet1': {'A1': 42, 'B2': 'x'}}
        write_cached(cache_dir, 'abc', result)
        assert read_cached(cache_dir, 'abc') == result


@test
def cache_miss_returns_none():
    with _tmp_path() as tmp_path:
        assert read_cached(tmp_path / 'cache', 'never') is None


@test
def cache_path_layout():
    with _tmp_path() as tmp_path:
        p = cache_path_for(tmp_path / 'cache', 'deadbeef')
        assert p == tmp_path / 'cache' / 'deadbeef.json'
