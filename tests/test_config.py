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
from sheetwright.security import SecurityLimits


@test
def project_config_round_trips():
    cfg = ProjectConfig(name='my-model', calc_engine='libreoffice')
    text = dump_project(cfg)
    assert load_project(text) == cfg


@test
def project_config_security_round_trips():
    sec = SecurityLimits(
        max_xlsx_uncompressed_bytes=10 * 1024 * 1024,
        max_xlsx_sheet_count=50,
        max_xlsx_cells_per_sheet=1_000_000,
        max_xlsx_shared_strings=2_000_000,
        soffice_timeout=60.0,
    )
    cfg = ProjectConfig(name='my-model', security=sec)
    text = dump_project(cfg)
    loaded = load_project(text)
    assert loaded.security == sec


@test
def project_config_missing_security_block_gives_none():
    toml = '[project]\nname = "x"\n'
    cfg = load_project(toml)
    assert cfg.security is None


@test
def project_config_partial_security_block_uses_defaults():
    toml = '[project]\nname = "x"\n[security]\nmax_xlsx_sheet_count = 10\n'
    cfg = load_project(toml)
    assert cfg.security is not None
    defaults = SecurityLimits.defaults()
    assert cfg.security.max_xlsx_sheet_count == 10
    assert (
        cfg.security.max_xlsx_uncompressed_bytes
        == defaults.max_xlsx_uncompressed_bytes
    )
    assert (
        cfg.security.max_xlsx_cells_per_sheet
        == defaults.max_xlsx_cells_per_sheet
    )
    assert (
        cfg.security.max_xlsx_shared_strings
        == defaults.max_xlsx_shared_strings
    )
    assert cfg.security.soffice_timeout == defaults.soffice_timeout


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
