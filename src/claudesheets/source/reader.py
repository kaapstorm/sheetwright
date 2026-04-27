"""Read a claudesheets project directory into a Workbook."""

from __future__ import annotations

from pathlib import Path

from claudesheets.config import load_project, load_workbook
from claudesheets.exceptions import ProjectError
from claudesheets.model.workbook import Sheet, Workbook
from claudesheets.project import slugify
from claudesheets.source.markdown import load_table
from claudesheets.source.yaml_sidecar import load_yaml


def read_source(project_dir: Path) -> Workbook:
    project_dir = Path(project_dir)
    load_project((project_dir / 'claudesheets.toml').read_text())
    manifest = load_workbook((project_dir / 'workbook.toml').read_text())

    wb = Workbook(name=manifest.name, named_ranges=list(manifest.named_ranges))
    for i, sheet_name in enumerate(manifest.sheets, start=1):
        stem = f'{i:02d}_{slugify(sheet_name)}'
        md_path = project_dir / 'sheets' / f'{stem}.md'
        yaml_path = project_dir / 'sheets' / f'{stem}.yaml'
        if not md_path.is_file():
            raise ProjectError(f'Missing sheet file: {md_path}')
        sheet = Sheet(name=sheet_name)
        load_table(sheet, md_path.read_text())
        if yaml_path.is_file():
            load_yaml(sheet, yaml_path.read_text())
        wb.sheets.append(sheet)
    return wb
