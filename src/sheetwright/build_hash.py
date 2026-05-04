"""Record and read the SHA-256 of the most recent built xlsx.

Used by the escape-hatch detection in `import`/`build`/`recalc` to
warn when the built xlsx has been edited externally.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Optional


@dataclass(frozen=True)
class BuildHashRecord:
    name: str
    sha256: str
    built_at: str  # ISO-8601 timestamp


def write_build_hash(path: Path, rec: BuildHashRecord) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(rec), indent=2, sort_keys=True))


def read_build_hash(path: Path) -> Optional[BuildHashRecord]:
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text())
    except json.JSONDecodeError:
        return None
    try:
        return BuildHashRecord(
            name=data['name'],
            sha256=data['sha256'],
            built_at=data['built_at'],
        )
    except (KeyError, TypeError):
        return None
