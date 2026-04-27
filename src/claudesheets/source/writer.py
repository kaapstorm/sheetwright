"""Write a Workbook to a claudesheets project directory."""

from __future__ import annotations

from pathlib import Path

from claudesheets.config import (
    ProjectConfig,
    WorkbookManifest,
    dump_project,
    dump_workbook,
)
from claudesheets.model.workbook import Workbook
from claudesheets.project import slugify
from claudesheets.source.markdown import dump_table
from claudesheets.source.yaml_sidecar import dump_yaml


def write_source(wb: Workbook, project_dir: Path) -> None:
    project_dir = Path(project_dir)
    project_dir.mkdir(parents=True, exist_ok=True)
    (project_dir / 'sheets').mkdir(exist_ok=True)
    (project_dir / 'data').mkdir(exist_ok=True)

    (project_dir / 'claudesheets.toml').write_text(
        dump_project(ProjectConfig(name=wb.name))
    )
    (project_dir / 'workbook.toml').write_text(
        dump_workbook(
            WorkbookManifest(
                name=wb.name,
                sheets=[s.name for s in wb.sheets],
                named_ranges=list(wb.named_ranges),
            )
        )
    )

    for i, sheet in enumerate(wb.sheets, start=1):
        stem = f'{i:02d}_{slugify(sheet.name)}'
        (project_dir / 'sheets' / f'{stem}.md').write_text(dump_table(sheet))
        sidecar = dump_yaml(sheet)
        if sidecar.strip():
            (project_dir / 'sheets' / f'{stem}.yaml').write_text(sidecar)
