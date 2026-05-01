"""Golden-file snapshots of calculated workbook values.

A `Snapshot` captures the calculated value of every formula cell in
a workbook. Literal-input cells are intentionally excluded.

Datetime values are normalized to ISO-8601 strings before snapshot
construction so JSON round-trips preserve equality.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

from claudesheets.calc.base import CalcResult
from claudesheets.model.cell import CellValue
from claudesheets.model.workbook import Workbook


@dataclass(frozen=True)
class Snapshot:
    values: Dict[str, Dict[str, CellValue]] = field(default_factory=dict)

    def write(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(self.values, indent=2, sort_keys=True))

    @classmethod
    def read(cls, path: Path) -> 'Snapshot':
        return cls(values=json.loads(path.read_text()))


def snapshot_from_calc_result(
    result: CalcResult, workbook: Workbook
) -> Snapshot:
    """Build a snapshot from a CalcResult, keeping only formula cells.

    `workbook` is consulted to determine which cells in the result
    were formulas in the source. Literal inputs are dropped.
    """
    formulas: Dict[str, set[str]] = {}
    for sheet in workbook.sheets:
        formulas[sheet.name] = {
            addr
            for addr, cell in sheet.cells.items()
            if cell.formula is not None
        }

    values: Dict[str, Dict[str, CellValue]] = {}
    for sheet_name, cells in result.items():
        keep = formulas.get(sheet_name, set())
        sheet_out: Dict[str, CellValue] = {}
        for addr, value in cells.items():
            if addr not in keep:
                continue
            sheet_out[addr] = _normalize(value)
        if sheet_out:
            values[sheet_name] = sheet_out
    return Snapshot(values=values)


def _normalize(value: CellValue) -> CellValue:
    if isinstance(value, datetime):
        return value.isoformat()
    return value


Diff = Tuple[str, str, CellValue, CellValue]


def diff_snapshots(a: Snapshot, b: Snapshot) -> List[Diff]:
    """Return (sheet, addr, old, new) tuples for every difference.

    `None` is used in either slot to signal "missing on that side".
    """
    diffs: List[Diff] = []
    sheets = sorted(set(a.values) | set(b.values))
    for s in sheets:
        addrs = sorted(set(a.values.get(s, {})) | set(b.values.get(s, {})))
        for addr in addrs:
            va = a.values.get(s, {}).get(addr)
            vb = b.values.get(s, {}).get(addr)
            if va != vb:
                diffs.append((s, addr, va, vb))
    return diffs
