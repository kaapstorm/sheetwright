from testsweet import test

from claudesheets.mcp.server import build_server, do_ping


@test
def build_server_returns_a_fastmcp_instance():
    server = build_server()
    assert server.name == 'claudesheets'


@test
def ping_tool_is_registered():
    server = build_server()
    tools = [t.name for t in server._tool_manager.list_tools()]
    assert 'do_ping' in tools


@test
def do_ping_returns_pong():
    assert do_ping() == 'pong'
