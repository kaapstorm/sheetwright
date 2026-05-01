from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.model.workbook import Workbook
from claudesheets.xlsx.reader import read_xlsx
from tests.fixtures.workbooks import (
    write_formatted_xlsx,
    write_simple_xlsx,
    write_validation_xlsx,
)


def _round_trip(tmp_path: Path, builder) -> Workbook:
    src = tmp_path / 'in.xlsx'
    builder(src)
    project = tmp_path / 'proj'
    project.mkdir()
    (project / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (project / 'workbook.toml').write_text(
        '[workbook]\nname = "in"\nsheets = []\n'
    )
    (project / 'sheets').mkdir()
    (project / 'data').mkdir()

    runner = CliRunner()
    r = runner.invoke(main, ['import', str(src), '--project', str(project)])
    assert r.exit_code == 0, r.output
    r = runner.invoke(main, ['build', '--project', str(project)])
    assert r.exit_code == 0, r.output
    return read_xlsx(project / 'build' / 'in.xlsx')


def _assert_workbooks_equivalent(a: Workbook, b: Workbook) -> None:
    assert [s.name for s in a.sheets] == [s.name for s in b.sheets]
    for sa, sb in zip(a.sheets, b.sheets):
        assert sa.cells.keys() == sb.cells.keys(), (
            f'cell sets differ on {sa.name}'
        )
        for addr in sa.cells:
            ca, cb = sa.cells[addr], sb.cells[addr]
            assert ca.value == cb.value, f'value differs at {sa.name}!{addr}'
            assert ca.formula == cb.formula, (
                f'formula differs at {sa.name}!{addr}'
            )
        assert sa.column_widths == sb.column_widths, (
            f'column widths differ on {sa.name}'
        )
        a_fmts = {
            ca.format_id and sa.formats.get(ca.format_id): addr
            for addr, ca in sa.cells.items()
            if ca.format_id
        }
        b_fmts = {
            cb.format_id and sb.formats.get(cb.format_id): addr
            for addr, cb in sb.cells.items()
            if cb.format_id
        }
        assert set(a_fmts.keys()) == set(b_fmts.keys()), (
            f'format set differs on {sa.name}'
        )
    assert {nr.name for nr in a.named_ranges} == {
        nr.name for nr in b.named_ranges
    }


@fixture
def tmp(tmp_path):
    yield tmp_path


@use(tmp)
def test_simple_workbook_round_trips_through_cli():
    out = _round_trip(tmp(), write_simple_xlsx)
    ref = tmp() / 'ref.xlsx'
    write_simple_xlsx(ref)
    src = read_xlsx(ref)
    _assert_workbooks_equivalent(src, out)


@use(tmp)
def test_formatted_workbook_round_trips_through_cli():
    out = _round_trip(tmp(), write_formatted_xlsx)
    ref = tmp() / 'ref.xlsx'
    write_formatted_xlsx(ref)
    src = read_xlsx(ref)
    _assert_workbooks_equivalent(src, out)


@use(tmp)
def test_validation_workbook_round_trips_through_cli():
    out = _round_trip(tmp(), write_validation_xlsx)
    ref = tmp() / 'ref.xlsx'
    write_validation_xlsx(ref)
    src = read_xlsx(ref)
    for sa, sb in zip(src.sheets, out.sheets):
        types_a = sorted(v.type for v in sa.validations)
        types_b = sorted(v.type for v in sb.validations)
        assert types_a == types_b
