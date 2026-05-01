import pytest

from claudesheets.model.workbook import NamedRange, Sheet, Workbook
from claudesheets.testing.addresses import parse_address


def _wb_with_named(name: str, ref: str) -> Workbook:
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    wb.named_ranges.append(NamedRange(name=name, ref=ref))
    return wb


def test_parse_qualified_a1():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'Inputs!B1') == ('Inputs', 'B1')


def test_parse_named_range_resolves_to_address():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'growth_rate') == ('Inputs', 'B1')


def test_parse_unknown_name_raises():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    with pytest.raises(KeyError, match='unknown'):
        parse_address(wb, 'no_such_name')


def test_parse_bare_a1_without_sheet_raises():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    with pytest.raises(ValueError, match='must include sheet'):
        parse_address(wb, 'B1')


def test_parse_strips_dollar_signs():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'Inputs!$B$1') == ('Inputs', 'B1')


def test_parse_quoted_sheet_name():
    wb = Workbook(name='x', sheets=[Sheet(name='My Sheet')])
    assert parse_address(wb, "'My Sheet'!A1") == ('My Sheet', 'A1')


def test_parse_sheet_scoped_named_range_resolves():
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    wb.named_ranges.append(
        NamedRange(
            name='local_rate',
            ref='Inputs!$B$5',
            scope='sheet',
            sheet='Inputs',
        )
    )
    # Sheet-scoped names resolve identically to workbook-scoped names
    # for our purposes: parse_address looks up by name and follows the
    # `ref`. Disambiguation between two sheet-scoped names with the
    # same string is out of scope for Plan 2 (no real workbook does
    # that).
    assert parse_address(wb, 'local_rate') == ('Inputs', 'B5')
