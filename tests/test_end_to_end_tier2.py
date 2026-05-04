import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from claudesheets.cli import main
from claudesheets.xlsx.reader import read_xlsx
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_tier2_xlsx


@contextmanager
def _project():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        write_tier2_xlsx(src)
        p = tmp_path / 'proj'
        p.mkdir()
        (p / 'claudesheets.toml').write_text(
            '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "in"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()
        runner = CliRunner()
        r = runner.invoke(main, ['import', str(src), '--project', str(p)])
        assert r.exit_code == 0, r.output
        r = runner.invoke(main, ['build', '--project', str(p)])
        assert r.exit_code == 0, r.output
        yield p


@test
def tier2_round_trips_through_cli():
    from claudesheets.model.conditional import (
        CellIsRule,
        ColorScaleRule,
    )

    with _project() as p:
        out = read_xlsx(p / 'build' / 'in.xlsx')
        s = out.sheet('S')

        assert s.frozen_panes == 'B2'
        assert s.print_area == 'A1:C3'
        assert 'B2' in s.comments
        assert s.comments['B2'].author == 'Alice'

        types = {type(cf) for cf in s.conditional_formats}
        assert CellIsRule in types
        assert ColorScaleRule in types

        assert len(s.tables) == 1
        assert s.tables[0].name == 'Sales'
        assert [c.name for c in s.tables[0].columns] == [
            'Region',
            'Q1',
            'Q2',
        ]


@test
@requires_libreoffice
def tier2_built_xlsx_evaluates_under_libreoffice():
    from claudesheets.calc.libreoffice import LibreOfficeEngine

    with _project() as p:
        built = p / 'build' / 'in.xlsx'
        result = LibreOfficeEngine().evaluate(built)
        assert any(result.values()), (
            f'libreoffice returned empty result for {built}; '
            'the built xlsx may be corrupt'
        )
