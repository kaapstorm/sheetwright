from pathlib import Path

from claudesheets.source.reader import read_source
from claudesheets.source.writer import write_source
from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_tier2_xlsx


def _project_dir(tmp_path: Path) -> Path:
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "x"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    return p


def test_tier2_features_round_trip_through_source(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_tier2_xlsx(src)
    project = _project_dir(tmp_path)

    wb_in = read_xlsx(src)
    write_source(wb_in, project)
    wb_via_src = read_source(project)

    s_in = wb_in.sheet('S')
    s_out = wb_via_src.sheet('S')

    assert s_out.frozen_panes == s_in.frozen_panes
    assert s_out.print_area == s_in.print_area
    assert s_out.comments == s_in.comments
    assert len(s_out.conditional_formats) == len(s_in.conditional_formats)
    assert {type(cf) for cf in s_out.conditional_formats} == {
        type(cf) for cf in s_in.conditional_formats
    }
    assert len(s_out.tables) == len(s_in.tables)
    assert s_out.tables[0].name == s_in.tables[0].name
