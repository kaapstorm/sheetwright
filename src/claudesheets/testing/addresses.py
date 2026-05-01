"""Resolve user-facing addresses to (sheet, A1) pairs.

Accepts:
- `Sheet!A1` or `'Sheet With Space'!A1` (Excel-style sheet-qualified)
- `name` (a workbook-scoped named range whose `ref` is a single cell)

Bare A1 (no sheet) is rejected: in this codebase tests always say
*which* sheet they mean.
"""

from __future__ import annotations

import re
from typing import Tuple

from claudesheets.model.workbook import Workbook


_QUOTED = re.compile(r"^'((?:[^']|'')+)'!(.+)$")
_UNQUOTED = re.compile(r'^([^!]+)!(.+)$')


def parse_address(wb: Workbook, address: str) -> Tuple[str, str]:
    if '!' in address:
        m = _QUOTED.match(address) or _UNQUOTED.match(address)
        if not m:
            raise ValueError(f'unparseable address: {address!r}')
        sheet, addr = m.group(1), m.group(2)
        sheet = sheet.replace("''", "'")
        return sheet, _strip_dollars(addr)

    for nr in wb.named_ranges:
        if nr.name == address:
            return parse_address(wb, nr.ref)

    if re.fullmatch(r'\$?[A-Za-z]+\$?\d+', address):
        raise ValueError(
            f'address {address!r} must include sheet (e.g. Sheet!A1)'
        )
    raise KeyError(f'unknown name: {address!r}')


def _strip_dollars(a1: str) -> str:
    return a1.replace('$', '')
