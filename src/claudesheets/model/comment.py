"""Cell comments (Excel "legacy" comments — author + text)."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Comment:
    author: str
    text: str
