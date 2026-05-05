import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from testsweet import catch_exceptions, test

from sheetwright.calc.libreoffice import LibreOfficeEngine, LibreOfficeError
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx
from sheetwright.security import SecurityLimits


@test
@requires_libreoffice
def libreoffice_engine_returns_calculated_values():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        eng = LibreOfficeEngine()
        out = eng.evaluate(src, limits=SecurityLimits.defaults())
        assert out['Inputs']['B1'] == 0.04
        assert out['Inputs']['B2'] == 1_000_000
        assert out['Outputs']['B1'] == 1_040_000


@test
@requires_libreoffice
def libreoffice_engine_propagates_string_values():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        out = LibreOfficeEngine().evaluate(
            src, limits=SecurityLimits.defaults()
        )
        assert out['Inputs']['A1'] == 'growth_rate'


@test
def libreoffice_engine_raises_when_binary_missing():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'fake.xlsx'
        src.write_bytes(b'not really an xlsx')
        eng = LibreOfficeEngine(soffice='/no/such/soffice/binary')
        with catch_exceptions() as excs:
            eng.evaluate(src, limits=SecurityLimits.defaults())
        assert excs and isinstance(excs[0], LibreOfficeError)
        assert 'not on' in str(excs[0])


@test
def libreoffice_engine_cmd_includes_hardening_flags():
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / 'fake.xlsx'
        src.write_bytes(b'fake')
        captured: list[list[str]] = []

        def fake_run(cmd, **kwargs):
            captured.append(list(cmd))
            raise FileNotFoundError('stub')

        with patch('sheetwright.calc.libreoffice.subprocess.run', fake_run):
            with catch_exceptions():
                LibreOfficeEngine().evaluate(
                    src, limits=SecurityLimits.defaults()
                )

        assert captured, 'subprocess.run was not called'
        cmd = captured[0]
        for flag in (
            '--safe-mode',
            '--norestore',
            '--nolockcheck',
            '--nofirststartwizard',
            '--nodefault',
        ):
            assert flag in cmd, f'{flag!r} missing from soffice cmd'


@test
def libreoffice_engine_tempdir_prefix():
    mock_td = MagicMock()
    mock_td.return_value.__enter__ = MagicMock(return_value='/tmp/stub')
    mock_td.return_value.__exit__ = MagicMock(return_value=False)

    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / 'fake.xlsx'
        src.write_bytes(b'fake')

    with patch(
        'sheetwright.calc.libreoffice.tempfile.TemporaryDirectory', mock_td
    ):
        with catch_exceptions():
            LibreOfficeEngine(soffice='/no/such/soffice').evaluate(
                src, limits=SecurityLimits.defaults()
            )

    assert mock_td.call_args is not None, 'TemporaryDirectory was not called'
    assert mock_td.call_args.kwargs.get('prefix') == 'sheetwright-calc-'
