"""Build the cached SQLite from data/*.csv."""

from __future__ import annotations

import csv
import re
import sqlite3
from pathlib import Path
from typing import List


def table_name_for(filename: str) -> str:
    stem = Path(filename).stem
    s = re.sub(r'[^A-Za-z0-9]+', '_', stem).strip('_').lower()
    return s or '_unnamed'


def _csvs_newer_than(csv_paths: List[Path], db_path: Path) -> bool:
    if not db_path.exists():
        return True
    db_mtime = db_path.stat().st_mtime
    return any(p.stat().st_mtime > db_mtime for p in csv_paths)


def build_bulk_cache(project_root: Path) -> Path:
    project_root = Path(project_root)
    data_dir = project_root / 'data'
    cache_dir = project_root / '.claudesheets'
    cache_dir.mkdir(parents=True, exist_ok=True)
    db_path = cache_dir / 'bulk.sqlite'

    csvs = sorted(data_dir.glob('*.csv')) if data_dir.is_dir() else []
    if not csvs:
        # No CSVs: ensure no stale db
        if db_path.exists():
            db_path.unlink()
        return db_path

    if not _csvs_newer_than(csvs, db_path):
        return db_path

    # Rebuild from scratch (deterministic).
    if db_path.exists():
        db_path.unlink()

    conn = sqlite3.connect(db_path)
    try:
        for csv_path in csvs:
            table = table_name_for(csv_path.name)
            with csv_path.open() as f:
                reader = csv.reader(f)
                header = next(reader, None)
                if not header:
                    continue
                # All columns typed TEXT in Plan 1; _schema.sql support
                # is deferred to a later plan.
                cols_sql = ', '.join(f'"{c}" TEXT' for c in header)
                conn.execute(f'CREATE TABLE "{table}" ({cols_sql})')
                placeholders = ', '.join(['?'] * len(header))
                conn.executemany(
                    f'INSERT INTO "{table}" VALUES ({placeholders})',
                    list(reader),
                )
        conn.commit()
    finally:
        conn.close()

    return db_path
