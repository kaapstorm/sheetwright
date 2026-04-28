import sqlite3
import time

from unmagic import fixture, use

from claudesheets.bulk import build_bulk_cache, table_name_for


@fixture
def project(tmp_path):
    proj = tmp_path / 'p'
    (proj / 'data').mkdir(parents=True)
    (proj / '.claudesheets').mkdir(parents=True)
    yield proj


def test_table_name_for_strips_extension():
    assert table_name_for('cpi_series.csv') == 'cpi_series'
    assert table_name_for('Panel-Data.csv') == 'panel_data'


@use(project)
def test_build_loads_csv_into_sqlite():
    p = project()
    (p / 'data' / 'series.csv').write_text(
        'year,value\n2020,1.0\n2021,2.0\n2022,3.0\n'
    )

    db_path = build_bulk_cache(p)
    assert db_path.exists()

    conn = sqlite3.connect(db_path)
    rows = conn.execute(
        'SELECT year, value FROM series ORDER BY year'
    ).fetchall()
    assert rows == [
        ('2020', '1.0'),
        ('2021', '2.0'),
        ('2022', '3.0'),
    ]


@use(project)
def test_build_skips_when_cache_is_newer():
    p = project()
    (p / 'data' / 's.csv').write_text('a,b\n1,2\n')
    db_path = build_bulk_cache(p)
    first_mtime = db_path.stat().st_mtime

    time.sleep(0.01)
    build_bulk_cache(p)
    second_mtime = db_path.stat().st_mtime
    assert first_mtime == second_mtime, 'cache rebuilt unnecessarily'


@use(project)
def test_build_rebuilds_when_csv_is_newer():
    p = project()
    (p / 'data' / 's.csv').write_text('a,b\n1,2\n')
    db_path = build_bulk_cache(p)
    first_mtime = db_path.stat().st_mtime

    time.sleep(0.05)
    (p / 'data' / 's.csv').write_text('a,b\n1,2\n3,4\n')
    build_bulk_cache(p)
    second_mtime = db_path.stat().st_mtime
    assert second_mtime > first_mtime
