import tempfile
from pathlib import Path

from testsweet import catch_exceptions, test

from claudesheets.calc.libreoffice import LibreOfficeEngine, LibreOfficeError
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@test
@requires_libreoffice
def libreoffice_engine_returns_calculated_values():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'in.xlsx'
        write_simple_xlsx(src)
        eng = LibreOfficeEngine()
        out = eng.evaluate(src)
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
        out = LibreOfficeEngine().evaluate(src)
        assert out['Inputs']['A1'] == 'growth_rate'


@test
def libreoffice_engine_raises_when_binary_missing():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        src = tmp_path / 'fake.xlsx'
        src.write_bytes(b'not really an xlsx')
        eng = LibreOfficeEngine(soffice='/no/such/soffice/binary')
        with catch_exceptions() as excs:
            eng.evaluate(src)
        assert excs and isinstance(excs[0], LibreOfficeError)
        assert 'not on' in str(excs[0])
