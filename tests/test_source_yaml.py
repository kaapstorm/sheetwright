from testsweet import test

from sheetwright.model.cell import Cell
from sheetwright.model.format import (
    Border,
    CellFormat,
    Fill,
    Font,
    Side,
)
from sheetwright.model.validation import DataValidation
from sheetwright.model.workbook import Sheet
from sheetwright.source.yaml_sidecar import dump_yaml, load_yaml


def _make_sheet() -> Sheet:
    sh = Sheet(name='S')
    sh.column_widths['A'] = 18.0
    sh.formats['fmt'] = CellFormat(
        font=Font(name='Calibri', size=11.0, bold=True, color='FF0000'),
        fill=Fill(color='FFFF00'),
        border=Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
        ),
        number_format='0.00%',
    )
    sh.set('A1', Cell(value=1, format_id='fmt'))
    sh.validations.append(
        DataValidation(
            type='list',
            ranges=['A1:A10'],
            formula1='"yes,no,maybe"',
            allow_blank=True,
        )
    )
    return sh


@test
def yaml_round_trips_column_widths():
    sh = _make_sheet()
    text = dump_yaml(sh)
    sh2 = Sheet(name='S')
    sh2.set('A1', Cell(value=1))
    load_yaml(sh2, text)
    assert sh2.column_widths['A'] == 18.0


@test
def yaml_round_trips_formats_and_cell_formats():
    sh = _make_sheet()
    text = dump_yaml(sh)
    sh2 = Sheet(name='S')
    sh2.set('A1', Cell(value=1))
    load_yaml(sh2, text)
    assert 'fmt' in sh2.formats
    fmt = sh2.formats['fmt']
    assert fmt.font.bold is True
    assert fmt.font.color == 'FF0000'
    assert fmt.fill.color == 'FFFF00'
    assert fmt.number_format == '0.00%'
    assert sh2.get('A1').format_id == 'fmt'


@test
def yaml_round_trips_validations():
    sh = _make_sheet()
    text = dump_yaml(sh)
    sh2 = Sheet(name='S')
    load_yaml(sh2, text)
    assert len(sh2.validations) == 1
    v = sh2.validations[0]
    assert v.type == 'list'
    assert v.ranges == ['A1:A10']
    assert v.formula1 == '"yes,no,maybe"'
