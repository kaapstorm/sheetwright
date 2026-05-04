import click
from testsweet import catch_exceptions, test

from sheetwright.mcp.errors import MCPError, classify_click_error


@test
def mcp_error_carries_code_and_message():
    e = MCPError('foo', 'something happened')
    assert e.code == 'foo'
    assert e.message == 'something happened'
    assert str(e) == 'something happened'


@test
def mcp_error_is_exception_subclass():
    assert issubclass(MCPError, Exception)


@test
def classify_external_refs():
    e = click.ClickException('Workbook contains external references; ...')
    assert classify_click_error(e) == 'external_refs'


@test
def classify_uncommitted():
    e = click.ClickException('You have uncommitted changes in sheets/.')
    assert classify_click_error(e) == 'uncommitted_source'


@test
def classify_project_not_found():
    e = click.ClickException('Not a sheetwright project: /nope')
    assert classify_click_error(e) == 'project_not_found'


@test
def classify_build_missing():
    e = click.ClickException('No built xlsx at /tmp/foo.xlsx.')
    assert classify_click_error(e) == 'build_missing'


@test
def classify_no_staged_session():
    e = click.ClickException(
        'No staged re-import session. Run `sheetwright import <xlsx> -I`.'
    )
    assert classify_click_error(e) == 'no_staged_session'


@test
def classify_no_staged_session_still_works():
    e = click.ClickException(
        'No staged re-import session. Run `sheetwright import <xlsx> -I` '
        'first.'
    )
    assert classify_click_error(e) == 'no_staged_session'


@test
def classify_staged_xlsx_missing():
    e = click.ClickException(
        'Staged xlsx no longer exists at /tmp/foo.xlsx. Re-stage with -I.'
    )
    assert classify_click_error(e) == 'staged_xlsx_changed'


@test
def classify_staged_xlsx_modified():
    e = click.ClickException(
        'Staged xlsx at /tmp/foo.xlsx has been modified since `-I` '
        '(recorded hash abc, current def). Re-stage with `-I`.'
    )
    assert classify_click_error(e) == 'staged_xlsx_changed'


@test
def classify_unknown_falls_back():
    e = click.ClickException('something else happened')
    assert classify_click_error(e) == 'click_error'


@test
def mcp_error_can_be_raised_and_caught():
    with catch_exceptions() as excs:
        raise MCPError('test_code', 'test message')
    assert excs and isinstance(excs[0], MCPError)
    assert excs[0].code == 'test_code'
