import sqlite3
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path

from testsweet import test

from claudesheets.bulk import build_bulk_cache, table_name_for


@contextmanager
def _project():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        proj = tmp_path / 'p'
        (proj / 'data').mkdir(parents=True)
        (proj / '.claudesheets').mkdir(parents=True)
        yield proj


@test
def table_name_for_strips_extension():
    assert table_name_for('cpi_series.csv') == 'cpi_series'
    assert table_name_for('Panel-Data.csv') == 'panel_data'


@test
def build_loads_csv_into_sqlite():
    with _project() as p:
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


@test
def build_skips_when_cache_is_newer():
    with _project() as p:
        (p / 'data' / 's.csv').write_text('a,b\n1,2\n')
        db_path = build_bulk_cache(p)
        first_mtime = db_path.stat().st_mtime

        time.sleep(0.01)
        build_bulk_cache(p)
        second_mtime = db_path.stat().st_mtime
        assert first_mtime == second_mtime, 'cache rebuilt unnecessarily'


@test
def build_rebuilds_when_csv_is_newer():
    with _project() as p:
        (p / 'data' / 's.csv').write_text('a,b\n1,2\n')
        db_path = build_bulk_cache(p)
        first_mtime = db_path.stat().st_mtime

        time.sleep(0.05)
        (p / 'data' / 's.csv').write_text('a,b\n1,2\n3,4\n')
        build_bulk_cache(p)
        second_mtime = db_path.stat().st_mtime
        assert second_mtime > first_mtime
