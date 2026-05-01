from pathlib import Path

import pytest

from claudesheets.calc import CalcEngine, CalcResult, get_calc_engine


def test_calc_engine_is_abstract():
    with pytest.raises(TypeError):
        CalcEngine()  # type: ignore[abstract]


def test_unknown_engine_raises():
    with pytest.raises(ValueError, match='unknown calc engine'):
        get_calc_engine('nonsense')


def test_calc_result_shape():
    result: CalcResult = {'Sheet1': {'A1': 1, 'B2': 'hello'}}
    assert result['Sheet1']['A1'] == 1


class _Recorder(CalcEngine):
    name = 'recorder'

    def evaluate(self, xlsx_path: Path) -> CalcResult:
        return {'Recorded': {'A1': str(xlsx_path)}}


def test_engine_subclass_evaluates(tmp_path):
    eng = _Recorder()
    out = eng.evaluate(tmp_path / 'x.xlsx')
    assert out == {'Recorded': {'A1': str(tmp_path / 'x.xlsx')}}
