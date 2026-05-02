"""Persisted re-import session state.

Bridges the `-I` non-interactive invocation (which records what would
happen) and the follow-up `--apply`/`--abort` invocation (which
either does it or discards). The state file lives at
`<project>/.claudesheets/reimport.json`.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Optional


@dataclass(frozen=True)
class ReimportSession:
    xlsx_path: str
    xlsx_sha256: str
    diff_summary: str
    created_at: str


def save_session(path: Path, session: ReimportSession) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(session), indent=2, sort_keys=True))


def load_session(path: Path) -> Optional[ReimportSession]:
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text())
        return ReimportSession(
            xlsx_path=data['xlsx_path'],
            xlsx_sha256=data['xlsx_sha256'],
            diff_summary=data['diff_summary'],
            created_at=data['created_at'],
        )
    except (json.JSONDecodeError, KeyError, TypeError):
        return None


def clear_session(path: Path) -> None:
    if path.is_file():
        path.unlink()
