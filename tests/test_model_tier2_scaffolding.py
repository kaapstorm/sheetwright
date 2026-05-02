import dataclasses

import pytest

from claudesheets.model.comment import Comment
from claudesheets.model.conditional import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
from claudesheets.model.table import ListTable, ListTableColumn
from claudesheets.model.workbook import Sheet


def test_sheet_defaults_for_tier2_fields():
    s = Sheet(name='S')
    assert s.print_area is None
    assert s.frozen_panes is None
    assert s.comments == {}
    assert s.conditional_formats == []
    assert s.tables == []


def test_comment_is_frozen():
    c = Comment(author='Alice', text='Watch this row')
    assert c.author == 'Alice'
    with pytest.raises(dataclasses.FrozenInstanceError):
        c.author = 'Bob'  # type: ignore[misc]


def test_cell_is_rule_typed_fields():
    cf = CellIsRule(
        ranges=('A1:A10',),
        operator='greaterThan',
        formula=('0',),
    )
    assert cf.operator == 'greaterThan'
    assert cf.formula == ('0',)
    with pytest.raises(dataclasses.FrozenInstanceError):
        cf.operator = 'lessThan'  # type: ignore[misc]


def test_formula_rule_typed_fields():
    cf = FormulaRule(ranges=('A1:A10',), formula='ISERROR(A1)')
    assert cf.formula == 'ISERROR(A1)'


def test_color_scale_rule_three_stops():
    cf = ColorScaleRule(
        ranges=('A1:A10',),
        start_type='min',
        start_color='FFFF0000',
        mid_type='percentile',
        mid_value='50',
        mid_color='FFFFFF00',
        end_type='max',
        end_color='FF00FF00',
    )
    assert cf.mid_type == 'percentile'


def test_data_bar_rule_typed_fields():
    cf = DataBarRule(ranges=('A1:A10',), color='FF638EC6')
    assert cf.color == 'FF638EC6'
    assert cf.show_value is True  # default


def test_icon_set_rule_typed_fields():
    cf = IconSetRule(ranges=('A1:A10',), icon_style='3TrafficLights1')
    assert cf.icon_style == '3TrafficLights1'
    assert cf.values == ('0', '33', '67')


def test_list_table_is_frozen():
    t = ListTable(
        name='Sales',
        ref='A1:C10',
        header_row_count=1,
        totals_row_count=0,
        columns=(
            ListTableColumn(name='Region'),
            ListTableColumn(name='Q1'),
            ListTableColumn(name='Q2'),
        ),
    )
    assert t.name == 'Sales'
    assert len(t.columns) == 3
    with pytest.raises(dataclasses.FrozenInstanceError):
        t.name = 'Other'  # type: ignore[misc]
