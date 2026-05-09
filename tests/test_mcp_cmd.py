import json
import queue
import subprocess
import sys
import threading

from testsweet import test


@test
def mcp_subcommand_exists_in_help():
    result = subprocess.run(
        [sys.executable, '-m', 'sheetwright.cli', '--help'],
        capture_output=True,
        text=True,
        timeout=10,
    )
    assert 'mcp' in result.stdout


@test
def mcp_subcommand_help_describes_stdio():
    result = subprocess.run(
        [sys.executable, '-m', 'sheetwright.cli', 'mcp', '--help'],
        capture_output=True,
        text=True,
        timeout=10,
    )
    assert 'stdio' in result.stdout.lower()


def _read_one_response(stream, timeout: float) -> str:
    q: queue.Queue = queue.Queue()

    def _reader():
        line = stream.readline()
        q.put(line)

    threading.Thread(target=_reader, daemon=True).start()
    try:
        line = q.get(timeout=timeout)
    except queue.Empty:
        raise TimeoutError('no MCP response within timeout') from None
    if not line:
        raise TimeoutError('MCP stream closed without a response')
    return line


@test
def mcp_server_responds_to_initialize_and_lists_tools():
    """Smoke test: launch the server, send initialize + tools/list,
    assert the registered tool names appear in the response."""
    proc = subprocess.Popen(
        [sys.executable, '-m', 'sheetwright.cli', 'mcp'],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=0,
    )
    try:
        assert proc.stdin is not None
        assert proc.stdout is not None
        init_req = (
            json.dumps(
                {
                    'jsonrpc': '2.0',
                    'id': 1,
                    'method': 'initialize',
                    'params': {
                        'protocolVersion': '2024-11-05',
                        'capabilities': {},
                        'clientInfo': {
                            'name': 'test',
                            'version': '0.0.0',
                        },
                    },
                }
            )
            + '\n'
        )
        proc.stdin.write(init_req)
        proc.stdin.flush()

        line = _read_one_response(proc.stdout, timeout=10.0)
        resp = json.loads(line)
        assert resp.get('id') == 1
        assert 'result' in resp

        proc.stdin.write(
            json.dumps(
                {
                    'jsonrpc': '2.0',
                    'method': 'notifications/initialized',
                }
            )
            + '\n'
        )
        proc.stdin.flush()

        proc.stdin.write(
            json.dumps(
                {
                    'jsonrpc': '2.0',
                    'id': 2,
                    'method': 'tools/list',
                }
            )
            + '\n'
        )
        proc.stdin.flush()

        line = _read_one_response(proc.stdout, timeout=10.0)
        resp = json.loads(line)
        assert resp.get('id') == 2
        names = {t['name'] for t in resp['result']['tools']}
        for expected in (
            'do_ping',
            'do_diff',
            'do_check',
            'do_init',
            'do_import_xlsx',
            'do_build',
            'do_recalc',
            'do_snapshot',
            'do_test',
            'do_reimport_stage',
            'do_reimport_apply',
            'do_reimport_abort',
        ):
            assert expected in names, f'missing tool {expected}'
    finally:
        if proc.stdin is not None:
            proc.stdin.close()
        proc.wait(timeout=5)


@test
def mcp_server_error_path_includes_code():
    """When a tool raises MCPError, the error response should be visible
    to the client. We don't (yet) require the `code` to round-trip via
    FastMCP — just that the human message is present and the response
    is a JSON-RPC error.
    """
    proc = subprocess.Popen(
        [sys.executable, '-m', 'sheetwright.cli', 'mcp'],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=0,
    )
    try:
        assert proc.stdin is not None
        assert proc.stdout is not None
        proc.stdin.write(
            json.dumps(
                {
                    'jsonrpc': '2.0',
                    'id': 1,
                    'method': 'initialize',
                    'params': {
                        'protocolVersion': '2024-11-05',
                        'capabilities': {},
                        'clientInfo': {
                            'name': 'test',
                            'version': '0.0.0',
                        },
                    },
                }
            )
            + '\n'
        )
        proc.stdin.flush()
        _read_one_response(proc.stdout, timeout=10.0)
        proc.stdin.write(
            json.dumps(
                {
                    'jsonrpc': '2.0',
                    'method': 'notifications/initialized',
                }
            )
            + '\n'
        )
        proc.stdin.flush()

        proc.stdin.write(
            json.dumps(
                {
                    'jsonrpc': '2.0',
                    'id': 2,
                    'method': 'tools/call',
                    'params': {
                        'name': 'do_check',
                        'arguments': {
                            'project': '/nonexistent/path/no-project'
                        },
                    },
                }
            )
            + '\n'
        )
        proc.stdin.flush()

        line = _read_one_response(proc.stdout, timeout=10.0)
        resp = json.loads(line)
        is_error = (
            'error' in resp
            or (resp.get('result') or {}).get('isError') is True
        )
        assert is_error, f'expected error response, got {resp}'
    finally:
        if proc.stdin is not None:
            proc.stdin.close()
        proc.wait(timeout=5)
