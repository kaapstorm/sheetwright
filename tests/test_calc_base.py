import tempfile
from pathlib import Path

from testsweet import catch_exceptions, test

from sheetwright.calc import CalcEngine, CalcResult, get_calc_engine


@test
def calc_engine_is_abstract():
    with catch_exceptions() as excs:
        CalcEngine()  # type: ignore[abstract]
    assert excs and isinstance(excs[0], TypeError)


@test
def unknown_engine_raises():
    with catch_exceptions() as excs:
        get_calc_engine('nonsense')
    assert excs and isinstance(excs[0], ValueError)
    assert 'unknown calc engine' in str(excs[0])


@test
def calc_result_shape():
    result: CalcResult = {'Sheet1': {'A1': 1, 'B2': 'hello'}}
    assert result['Sheet1']['A1'] == 1


class _Recorder(CalcEngine):
    def evaluate(self, xlsx_path: Path) -> CalcResult:
        return {'Recorded': {'A1': str(xlsx_path)}}


@test
def engine_subclass_evaluates():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        eng = _Recorder()
        out = eng.evaluate(tmp_path / 'x.xlsx')
        assert out == {'Recorded': {'A1': str(tmp_path / 'x.xlsx')}}
