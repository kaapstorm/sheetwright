from claudesheets.mcp.server import build_server, do_ping


def test_build_server_returns_a_fastmcp_instance():
    server = build_server()
    assert server.name == 'claudesheets'


def test_ping_tool_is_registered():
    server = build_server()
    tools = [t.name for t in server._tool_manager.list_tools()]
    assert 'do_ping' in tools


def test_do_ping_returns_pong():
    assert do_ping() == 'pong'
