import click
import pytest

from claudesheets.mcp.errors import MCPError, classify_click_error


def test_mcp_error_carries_code_and_message():
    e = MCPError('foo', 'something happened')
    assert e.code == 'foo'
    assert e.message == 'something happened'
    assert str(e) == 'something happened'


def test_mcp_error_is_exception_subclass():
    assert issubclass(MCPError, Exception)


def test_classify_external_refs():
    e = click.ClickException('Workbook contains external references; ...')
    assert classify_click_error(e) == 'external_refs'


def test_classify_uncommitted():
    e = click.ClickException('You have uncommitted changes in sheets/.')
    assert classify_click_error(e) == 'uncommitted_source'


def test_classify_project_not_found():
    e = click.ClickException('Not a claudesheets project: /nope')
    assert classify_click_error(e) == 'project_not_found'


def test_classify_build_missing():
    e = click.ClickException('No built xlsx at /tmp/foo.xlsx.')
    assert classify_click_error(e) == 'build_missing'


def test_classify_no_staged_session():
    e = click.ClickException(
        'No staged re-import session. Run `claudesheets import <xlsx> -I`.'
    )
    assert classify_click_error(e) == 'no_staged_session'


def test_classify_unknown_falls_back():
    e = click.ClickException('something else happened')
    assert classify_click_error(e) == 'click_error'


def test_mcp_error_can_be_raised_and_caught():
    with pytest.raises(MCPError) as excinfo:
        raise MCPError('test_code', 'test message')
    assert excinfo.value.code == 'test_code'
