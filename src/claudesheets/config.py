"""Read/write claudesheets.toml and workbook.toml."""

from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from typing import Any, Dict, List

import tomli_w

from claudesheets.model.workbook import NamedRange


@dataclass(frozen=True)
class ProjectConfig:
    name: str
    calc_engine: str = 'libreoffice'


@dataclass
class WorkbookManifest:
    name: str
    sheets: List[str] = field(default_factory=list)
    named_ranges: List[NamedRange] = field(default_factory=list)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, WorkbookManifest):
            return NotImplemented
        return (
            self.name == other.name
            and self.sheets == other.sheets
            and len(self.named_ranges) == len(other.named_ranges)
            and all(
                a.name == b.name
                and a.scope == b.scope
                and a.sheet == b.sheet
                and a.ref == b.ref
                for a, b in zip(self.named_ranges, other.named_ranges)
            )
        )


def dump_project(cfg: ProjectConfig) -> str:
    return tomli_w.dumps(
        {
            'project': {'name': cfg.name},
            'build': {'calc_engine': cfg.calc_engine},
        }
    )


def load_project(text: str) -> ProjectConfig:
    data = tomllib.loads(text)
    return ProjectConfig(
        name=data['project']['name'],
        calc_engine=data.get('build', {}).get('calc_engine', 'libreoffice'),
    )


def dump_workbook(m: WorkbookManifest) -> str:
    nrs = [
        {
            'name': nr.name,
            'scope': nr.scope,
            **({'sheet': nr.sheet} if nr.sheet else {}),
            'ref': nr.ref,
        }
        for nr in m.named_ranges
    ]
    doc: Dict[str, Any] = {
        'workbook': {'name': m.name, 'sheets': list(m.sheets)},
    }
    if nrs:
        doc['named_ranges'] = nrs
    return tomli_w.dumps(doc)


def load_workbook(text: str) -> WorkbookManifest:
    data = tomllib.loads(text)
    nrs = [
        NamedRange(
            name=d['name'],
            ref=d['ref'],
            scope=d.get('scope', 'workbook'),
            sheet=d.get('sheet'),
        )
        for d in data.get('named_ranges', [])
    ]
    wb = data['workbook']
    return WorkbookManifest(
        name=wb['name'],
        sheets=list(wb.get('sheets', [])),
        named_ranges=nrs,
    )
