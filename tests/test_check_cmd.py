import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main


@contextmanager
def _project():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        p = tmp_path / 'proj'
        p.mkdir()
        (p / 'sheetwright.toml').write_text(
            '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
        )
        (p / 'workbook.toml').write_text(
            '[workbook]\nname = "x"\nsheets = []\n'
        )
        (p / 'sheets').mkdir()
        (p / 'data').mkdir()
        yield p


@test
def check_clean_returns_zero():
    with _project() as p:
        runner = CliRunner()
        r = runner.invoke(main, ['check', '--project', str(p)])
        assert r.exit_code == 0
        assert 'no issues' in r.output.lower()


@test
def check_finds_orphaned_sheet():
    with _project() as p:
        (p / 'sheets' / '01_orphan.md').write_text(
            '| (cell) | A |\n| --- | --- |\n'
        )
        runner = CliRunner()
        r = runner.invoke(main, ['check', '--project', str(p)])
        assert r.exit_code != 0
        assert 'sheet_file_missing_from_manifest' in r.output
