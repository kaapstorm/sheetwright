import tempfile
from contextlib import contextmanager
from pathlib import Path

from click.testing import CliRunner
from testsweet import test

from claudesheets.cli import main
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@contextmanager
def _built():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
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
        runner.invoke(main, ['import', str(src), '--project', str(p)])
        runner.invoke(main, ['build', '--project', str(p)])
        yield tmp_path, p


@test
def no_warning_when_build_unchanged():
    with _built() as (_tmp, p):
        runner = CliRunner()
        r = runner.invoke(main, ['snapshot', '--project', str(p)])
        assert 'modified externally' not in r.output


@test
@requires_libreoffice
def warn_when_build_changed_externally():
    with _built() as (_tmp, p):
        built_xlsx = p / 'build' / 'in.xlsx'
        data = built_xlsx.read_bytes()
        built_xlsx.write_bytes(data + b' ')
        runner = CliRunner()
        r = runner.invoke(main, ['snapshot', '--project', str(p)])
        assert 'modified externally' in r.output.lower()


@test
def no_warning_when_no_record_yet():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        p = tmp_path / 'fresh'
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
        r = runner.invoke(main, ['snapshot', '--project', str(p)])
        assert 'modified externally' not in r.output


@test
def warn_when_workbook_renamed_orphans_hash():
    with _built() as (_tmp, p):
        cs_toml = p / 'claudesheets.toml'
        cs_toml.write_text(
            cs_toml.read_text().replace('name = "in"', 'name = "renamed"')
        )
        runner = CliRunner()
        r = runner.invoke(main, ['snapshot', '--project', str(p)])
        assert 'orphaned' in r.output.lower() or 'orphan' in r.output.lower()
