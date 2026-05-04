import tempfile
from pathlib import Path

from testsweet import test

from claudesheets.model.cell import Cell
from claudesheets.model.format import CellFormat, Font
from claudesheets.model.validation import DataValidation
from claudesheets.model.workbook import NamedRange, Sheet, Workbook
from claudesheets.source.reader import read_source
from claudesheets.source.writer import write_source


def _make_workbook() -> Workbook:
    wb = Workbook(name='m')
    inputs = Sheet(name='Inputs')
    inputs.set('A1', Cell(value='growth'))
    inputs.set('B1', Cell(value=0.04))
    inputs.formats['bold'] = CellFormat(
        font=Font(name='Calibri', size=11.0, bold=True)
    )
    inputs.set('A1', Cell(value='growth', format_id='bold'))
    inputs.column_widths['A'] = 18.0
    inputs.validations.append(
        DataValidation(
            type='list',
            ranges=['A1:A10'],
            formula1='"yes,no,maybe"',
            allow_blank=True,
        )
    )
    outputs = Sheet(name='Outputs')
    outputs.set('A1', Cell(value='rev'))
    outputs.set('B1', Cell(formula='=Inputs!B1*100'))
    wb.sheets = [inputs, outputs]
    wb.named_ranges.append(
        NamedRange(name='growth_rate', ref='Inputs!$B$1', scope='workbook')
    )
    return wb


@test
def round_trip_via_source_dir():
    with tempfile.TemporaryDirectory() as td:
        project_dir = Path(td) / 'proj'
        wb = _make_workbook()
        write_source(wb, project_dir)
        wb2 = read_source(project_dir)
        assert [s.name for s in wb2.sheets] == ['Inputs', 'Outputs']
        assert wb2.sheet('Inputs').get('A1').value == 'growth'
        assert wb2.sheet('Inputs').get('A1').format_id == 'bold'
        assert wb2.sheet('Inputs').column_widths['A'] == 18.0
        assert wb2.sheet('Outputs').get('B1').formula == '=Inputs!B1*100'
        assert wb2.named_ranges[0].name == 'growth_rate'
        assert wb2.named_ranges[0].ref == 'Inputs!$B$1'
        v = wb2.sheet('Inputs').validations[0]
        assert v.ranges == ['A1:A10']
        assert v.formula1 == '"yes,no,maybe"'
