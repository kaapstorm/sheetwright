from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def baseline(tmp_path: Path):
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
    runner = CliRunner()
    runner.invoke(main, ['import', str(src), '--project', str(p)])
    runner.invoke(main, ['build', '--project', str(p)])
    yield p, src


@use(baseline, requires_libreoffice)
def test_full_escape_hatch_loop():
    p, _orig_src = baseline()
    runner = CliRunner()

    # External edit: tamper with build/<name>.xlsx by importing a
    # different value via a separate xlsx.
    edited = p / 'build' / 'in.xlsx'
    import openpyxl

    wb = openpyxl.load_workbook(edited)
    wb['Inputs']['B1'] = 0.10  # was 0.04
    wb.save(edited)

    # snapshot should now warn that the build is externally edited.
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert 'modified externally' in r.output.lower()

    # User runs `import build/in.xlsx -I` to stage the diff.
    r = runner.invoke(
        main,
        ['import', str(edited), '-I', '--project', str(p)],
    )
    assert r.exit_code == 0
    assert 'Inputs!B1' in r.output
    assert (p / '.claudesheets' / 'reimport.json').is_file()

    # Apply the staged changes.
    r = runner.invoke(main, ['import', '--apply', '--project', str(p)])
    assert r.exit_code == 0
    md = p / 'sheets' / '01_inputs.md'
    assert '0.1' in md.read_text()
