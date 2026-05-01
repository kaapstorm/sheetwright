from pathlib import Path

import pytest
from unmagic import use

from claudesheets.calc.libreoffice import LibreOfficeEngine, LibreOfficeError
from tests.fixtures.libreoffice import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@use(requires_libreoffice)
def test_libreoffice_engine_returns_calculated_values(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    eng = LibreOfficeEngine()
    out = eng.evaluate(src)
    # Inputs!B1 = 0.04 (literal), Inputs!B2 = 1_000_000 (literal),
    # Outputs!B1 = =Inputs!B2 * (1 + Inputs!B1) = 1_040_000.
    assert out['Inputs']['B1'] == 0.04
    assert out['Inputs']['B2'] == 1_000_000
    assert out['Outputs']['B1'] == 1_040_000


@use(requires_libreoffice)
def test_libreoffice_engine_propagates_string_values(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    out = LibreOfficeEngine().evaluate(src)
    assert out['Inputs']['A1'] == 'growth_rate'


def test_libreoffice_engine_raises_when_binary_missing(tmp_path: Path):
    # No need for soffice on $PATH — we point at a path that doesn't exist.
    src = tmp_path / 'fake.xlsx'
    src.write_bytes(b'not really an xlsx')
    eng = LibreOfficeEngine(soffice='/no/such/soffice/binary')
    with pytest.raises(LibreOfficeError, match='not on'):
        eng.evaluate(src)
