from testsweet import test

from sheetwright.config import (
    ProjectConfig,
    WorkbookManifest,
    dump_project,
    dump_workbook,
    load_project,
    load_workbook,
)
from sheetwright.model.workbook import NamedRange


@test
def project_config_round_trips():
    cfg = ProjectConfig(name='my-model', calc_engine='libreoffice')
    text = dump_project(cfg)
    assert load_project(text) == cfg


@test
def workbook_manifest_round_trips_basic():
    m = WorkbookManifest(
        name='my-model', sheets=['Inputs', 'Outputs'], named_ranges=[]
    )
    assert load_workbook(dump_workbook(m)) == m


@test
def workbook_manifest_round_trips_named_ranges():
    m = WorkbookManifest(
        name='m',
        sheets=['S'],
        named_ranges=[
            NamedRange(name='x', scope='workbook', ref='S!$A$1'),
            NamedRange(name='y', scope='sheet', sheet='S', ref='$B$1'),
        ],
    )
    assert load_workbook(dump_workbook(m)) == m
