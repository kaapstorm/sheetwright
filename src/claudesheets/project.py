"""Locate and validate a claudesheets project directory."""

from __future__ import annotations

import re
from pathlib import Path

from claudesheets.exceptions import ProjectError


class Project:
    def __init__(self, root: Path):
        self.root = Path(root).resolve()

    @classmethod
    def open(cls, path: str | Path) -> 'Project':
        root = Path(path).resolve()
        if not (root / 'claudesheets.toml').is_file():
            raise ProjectError(f'Not a claudesheets project: {root}')
        return cls(root)

    @property
    def claudesheets_toml(self) -> Path:
        return self.root / 'claudesheets.toml'

    @property
    def workbook_toml(self) -> Path:
        return self.root / 'workbook.toml'

    @property
    def sheets_dir(self) -> Path:
        return self.root / 'sheets'

    @property
    def data_dir(self) -> Path:
        return self.root / 'data'

    @property
    def build_dir(self) -> Path:
        return self.root / 'build'

    @property
    def cache_dir(self) -> Path:
        return self.root / '.claudesheets'

    @property
    def imports_dir(self) -> Path:
        return self.root / 'imports'


def slugify(name: str) -> str:
    s = re.sub(r'[^A-Za-z0-9]+', '_', name).strip('_').lower()
    return s or 'sheet'
