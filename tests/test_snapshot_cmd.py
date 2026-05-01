from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def built(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text('[workbook]\nname = "in"\nsheets = []\n')
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    (p / 'tests').mkdir()
    runner = CliRunner()
    runner.invoke(main, ['import', str(src), '--project', str(p)])
    runner.invoke(main, ['build', '--project', str(p)])
    yield p


@use(built, requires_libreoffice)
def test_snapshot_initializes_when_missing():
    p = built()
    runner = CliRunner()
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code == 0
    assert (p / 'tests' / 'snapshots' / 'in.json').is_file()
    assert 'initialized' in r.output.lower()


@use(built, requires_libreoffice)
def test_snapshot_clean_when_unchanged():
    p = built()
    runner = CliRunner()
    runner.invoke(main, ['snapshot', '--project', str(p)])
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code == 0
    assert 'no changes' in r.output.lower()


@use(built, requires_libreoffice)
def test_snapshot_reports_diff_and_exits_nonzero():
    import json

    p = built()
    runner = CliRunner()
    runner.invoke(main, ['snapshot', '--project', str(p)])
    snap = p / 'tests' / 'snapshots' / 'in.json'
    data = json.loads(snap.read_text())
    data['Outputs']['B1'] = 999_999
    snap.write_text(json.dumps(data, indent=2, sort_keys=True))
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code != 0
    assert '999999' in r.output or 'Outputs!B1' in r.output


@use(built, requires_libreoffice)
def test_snapshot_update_overwrites():
    p = built()
    runner = CliRunner()
    runner.invoke(main, ['snapshot', '--project', str(p)])
    snap = p / 'tests' / 'snapshots' / 'in.json'
    snap.write_text('{}')
    r = runner.invoke(main, ['snapshot', '--project', str(p), '--update'])
    assert r.exit_code == 0
    text = snap.read_text()
    assert '"Outputs"' in text
