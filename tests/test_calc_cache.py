from pathlib import Path

from claudesheets.calc.cache import (
    cache_path_for,
    hash_xlsx,
    read_cached,
    write_cached,
)


def test_hash_xlsx_is_stable(tmp_path: Path):
    p = tmp_path / 'a.xlsx'
    p.write_bytes(b'hello')
    h1 = hash_xlsx(p)
    h2 = hash_xlsx(p)
    assert h1 == h2
    assert len(h1) == 64  # sha256 hex


def test_hash_xlsx_changes_with_content(tmp_path: Path):
    a = tmp_path / 'a.xlsx'
    b = tmp_path / 'b.xlsx'
    a.write_bytes(b'hello')
    b.write_bytes(b'world')
    assert hash_xlsx(a) != hash_xlsx(b)


def test_cache_round_trips(tmp_path: Path):
    cache_dir = tmp_path / 'cache'
    result = {'Sheet1': {'A1': 42, 'B2': 'x'}}
    write_cached(cache_dir, 'abc', result)
    assert read_cached(cache_dir, 'abc') == result


def test_cache_miss_returns_none(tmp_path: Path):
    assert read_cached(tmp_path / 'cache', 'never') is None


def test_cache_path_layout(tmp_path: Path):
    p = cache_path_for(tmp_path / 'cache', 'deadbeef')
    assert p == tmp_path / 'cache' / 'deadbeef.json'
