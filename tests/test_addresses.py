from testsweet import catch_exceptions, test

from sheetwright.model.workbook import NamedRange, Sheet, Workbook
from sheetwright.testing.addresses import parse_address


def _wb_with_named(name: str, ref: str) -> Workbook:
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    wb.named_ranges.append(NamedRange(name=name, ref=ref))
    return wb


@test
def parse_qualified_a1():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'Inputs!B1') == ('Inputs', 'B1')


@test
def parse_named_range_resolves_to_address():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'growth_rate') == ('Inputs', 'B1')


@test
def parse_unknown_name_raises():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    with catch_exceptions() as excs:
        parse_address(wb, 'no_such_name')
    assert excs and isinstance(excs[0], KeyError)
    assert 'unknown' in str(excs[0])


@test
def parse_bare_a1_without_sheet_raises():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    with catch_exceptions() as excs:
        parse_address(wb, 'B1')
    assert excs and isinstance(excs[0], ValueError)
    assert 'must include sheet' in str(excs[0])


@test
def parse_strips_dollar_signs():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'Inputs!$B$1') == ('Inputs', 'B1')


@test
def parse_quoted_sheet_name():
    wb = Workbook(name='x', sheets=[Sheet(name='My Sheet')])
    assert parse_address(wb, "'My Sheet'!A1") == ('My Sheet', 'A1')


@test
def parse_sheet_scoped_named_range_resolves():
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
