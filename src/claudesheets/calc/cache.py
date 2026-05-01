"""Content-addressed calc cache.

The cache key is the SHA-256 of the built xlsx bytes. Hits avoid
running the calc engine entirely.
"""

from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Optional

from claudesheets.calc.base import CalcResult


def hash_xlsx(xlsx_path: Path) -> str:
    h = hashlib.sha256()
    with open(xlsx_path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def cache_path_for(cache_dir: Path, key: str) -> Path:
    return cache_dir / f'{key}.json'


def read_cached(cache_dir: Path, key: str) -> Optional[CalcResult]:
    p = cache_path_for(cache_dir, key)
    if not p.is_file():
        return None
    data = json.loads(p.read_text())
    return data['result']


def write_cached(cache_dir: Path, key: str, result: CalcResult) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    p = cache_path_for(cache_dir, key)
    p.write_text(
        json.dumps(
            {
                'key': key,
                'computed_at': datetime.utcnow().isoformat() + 'Z',
                'result': result,
            },
            indent=2,
            sort_keys=True,
            default=_json_default,
        )
    )
    return p


def _json_default(obj: object) -> object:
    if isinstance(obj, datetime):
        return obj.isoformat() + 'Z'
    raise TypeError(f'not JSON serializable: {type(obj).__name__}')
