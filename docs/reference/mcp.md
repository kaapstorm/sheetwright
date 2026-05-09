# MCP server

sheetwright ships an MCP (Model Context Protocol) server that
exposes every CLI command as a typed tool. Launch it with
`sheetwright mcp`; the server runs over stdio and stays alive until
the client closes the connection.

## Wiring

Most MCP clients launch `sheetwright mcp` as a subprocess and route
MCP traffic over stdin/stdout. For Claude Desktop, edit the config
file for your platform:

- macOS:
  `~/Library/Application Support/Claude/claude_desktop_config.json`
- Linux: `~/.config/Claude/claude_desktop_config.json`
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

Add an entry like:

```json
{
  "mcpServers": {
    "sheetwright": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/your/model", "sheetwright", "mcp"]
    }
  }
}
```

On Windows, use a full Windows path (and double the backslashes in
JSON), for example
`"C:\\Users\\you\\my-model"`.

For a generic stdio MCP client:

```bash
sheetwright mcp
```

Note: `--directory` is `uv`'s flag — it sets the subprocess's working
directory before invoking `sheetwright mcp`. It is not a sheetwright
workspace root. The server is per-call sandboxed via the `project`
argument of each tool, not via the launch cwd.

The tools all accept a `project` argument (path to a sheetwright
project) so a single server can drive multiple projects sequentially.

## Trust model and path containment

sheetwright uses a per-call workspace root. Every tool call carries a
`project` argument that is the sandbox for that call: every other path
argument on the same call must resolve under that `project` directory.
Three carve-outs exist for paths that are intentionally outside the
project tree:

- `do_init(path)` — `path` is the directory to scaffold; it need not
  exist yet, so containment is not enforced.
- `do_import_xlsx(xlsx, ...)` — the source xlsx may live anywhere the
  operator's user can read; only the destination `project` is
  sandboxed.
- `do_reimport_stage(xlsx, ...)` — same as above; the xlsx being
  re-imported is an external file.

The "one server, many projects" property is preserved: the `project`
argument changes per call, and the server enforces containment relative
to whichever root was passed on that call.

The server does not maintain an allow-list of permitted project
directories. A client may pass any `project` path that the operator's
user has read access to. The operator is the gatekeeper — they control
which projects are reachable by deciding which user runs the server and
what filesystem access that user has.

`do_test` runs test files found under `<project>/tests/` in-process via
`exec_module`. This is intentional: the operator chose to open that
project, which implies trusting its test code. Test discovery is
restricted to the `tests/` subdirectory and cannot escape the project
root.

xlsx ingest enforces size, sheet-count, and cell-count limits at two
layers. The operator sets a ceiling via environment variables
(`SHEETWRIGHT_MAX_XLSX_BYTES` etc.). A project's `sheetwright.toml` may
tighten those limits further via a `[security]` block but cannot raise
them above the operator ceiling. LibreOffice (used for recalc) runs
`--safe-mode` with an isolated profile directory.

**Known gaps (Phase 1):**

- Parser CPU time is not bounded. An xlsx within the byte cap can still
  consume some seconds of CPU during parsing.
- Path validation (`resolve_under`) has a TOCTOU window between the
  check and openpyxl's open. Phase 1 assumes the filesystem is stable
  during a tool call.

## Tools

### `do_ping()`

Smoke test. Returns the literal string `"pong"`. Useful for
confirming the server is alive.

### `do_init(path: str)`

Scaffold an empty sheetwright project at `path`.

Returns: `{ok: bool, message: str}`.

Errors: `click_error` (e.g. directory non-empty).

### `do_import_xlsx(xlsx, project, archive=False, flatten=False)`

Initial import of an `.xlsx` into source form.

| Arg | Type | Description |
| --- | --- | --- |
| `xlsx` | str | Path to the xlsx. |
| `project` | str | Project directory. |
| `archive` | bool | Copy the xlsx into `imports/`. |
| `flatten` | bool | Replace external-reference formulas with cached values. |

Returns: `{ok: bool, message: str}`.

Errors: `project_not_found`, `external_refs` (workbook contains
external refs and `flatten=False`), `reimport_required` (`sheets/`
is non-empty — use the re-import tools instead), `click_error`.

### `do_build(project, out_path=None)`

Compile sources into an `.xlsx`.

| Arg | Type | Description |
| --- | --- | --- |
| `project` | str | Project directory. |
| `out_path` | str \| None | Output path; defaults to `build/<name>.xlsx`. |

Returns: `{ok: bool, message: str}`.

Errors: `project_not_found`, `click_error`.

### `do_recalc(project, force=False)`

Run the calc engine and cache results.

Returns: `{ok: bool, message: str}`.

Errors: `project_not_found`, `build_missing` (no `build/<n>.xlsx`),
`click_error`.

### `do_snapshot(project, update=False)`

Compare or update the snapshot of calculated values.

Returns: `{ok: bool, message: str, has_diffs: bool}`.

`has_diffs=True` is **not** an error — it's a successful tool call
that surfaced a difference. The client decides how to react.

Errors: `project_not_found`, `build_missing`, `click_error`.

### `do_test(project, targets=[])`

Run the project's testsweet tests.

| Arg | Type | Description |
| --- | --- | --- |
| `project` | str | Project directory. |
| `targets` | list[str] | Specific test files / dirs; `[]` for default discovery. |

Returns: `{passed: bool, output: str}`. `passed=False` is a
successful tool call with failing tests, not a tool error.

Errors: `project_not_found`, `click_error`.

### `do_diff(project, vs=None)`

Diff source against a target workbook.

| Arg | Type | Description |
| --- | --- | --- |
| `project` | str | Project directory. |
| `vs` | str \| None | `"xlsx:<path>"` or `"source:<path>"`; `None` defaults to `build/<name>.xlsx`. |

Returns: `{is_empty: bool, rendered: str, structured: dict}`.

`rendered` is the human-readable diff; `structured` is the
machine-readable form (cells added, removed, changed, with addresses
and old/new values).

Errors: `project_not_found`, `click_error`.

### `do_check(project)`

Lint dangling references and schema mismatches.

Returns: `{issues: [{kind: str, detail: str, location: str | null}, ...]}`.

Issue kinds match the [CLI reference](cli.md#check):
`dangling_sheet_ref`, `dangling_named_range`, `manifest_sheet_missing`,
`sheet_file_missing_from_manifest`, `bulk_data_missing`.

Errors: `project_not_found`, `click_error`.

### `do_reimport_stage(xlsx, project, flatten=False, force=False)`

Compute and stage a re-import diff. If the diff is non-empty, saves
a session under `.sheetwright/reimport.json` that
`do_reimport_apply` can later commit.

Returns:

```
{
  is_empty: bool,
  rendered: str,
  structured: {...},
  rendered_diff: str,
  xlsx_path: str,
  xlsx_sha256: str,
}
```

`xlsx_path` and `xlsx_sha256` let the client verify between calls
that the staged xlsx hasn't changed.

Errors: `project_not_found`, `external_refs`, `uncommitted_source`
(use `force=True` to skip), `click_error`.

### `do_reimport_apply(project, archive=False)`

Complete a previously-staged re-import.

Returns: `{ok: bool, message: str}`.

Errors: `project_not_found`, `no_staged_session`,
`staged_xlsx_changed` (the staged xlsx's hash no longer matches —
re-stage), `click_error`.

### `do_reimport_abort(project)`

Discard a previously-staged re-import session.

Returns: `{ok: bool, message: str}`.

Errors: `project_not_found`.

## Error codes

All MCP tool errors raise `MCPError(code, message)`. The codes are
stable; clients can branch on them. The full list:

| Code | Meaning |
| --- | --- |
| `project_not_found` | The path is not a sheetwright project (no `sheetwright.toml`). |
| `external_refs` | The xlsx has cross-workbook references and `flatten=False`. |
| `uncommitted_source` | `sheets/` has uncommitted git changes; pass `force=True` to skip. |
| `build_missing` | No `build/<name>.xlsx`; run `do_build` first. |
| `no_staged_session` | `do_reimport_apply` / `do_reimport_abort` called without a staged session. |
| `staged_xlsx_changed` | The xlsx pointed at by the staged session has been modified since `do_reimport_stage`. |
| `reimport_required` | `do_import_xlsx` against a populated source dir; use `do_reimport_*` instead. |
| `click_error` | Unrecognised CLI error; check `message`. |

The codes are defined in `src/sheetwright/mcp/errors.py`. Adding a
new one is a one-line change in `classify_click_error` once a real
client wants to branch on it.

## Lifecycle

The server exposes 12 tools (one per command, plus the three
re-import operations and `do_ping`). Tool definitions live in
`src/sheetwright/mcp/server.py`. The server is built with
`fastmcp` and is a thin facade over `commands/<name>.run` and the
pure helpers in `sheetwright.diff` / `sheetwright.reimport`.

A typical Claude-driven session looks like:

1. `do_init` (or `do_import_xlsx`) — scaffold or onboard.
2. Edit source via the filesystem (Claude does this directly).
3. `do_build`, `do_check`, `do_diff` — verify.
4. `do_recalc`, `do_test`, `do_snapshot` — assert behaviour.
5. On regression: `do_reimport_stage` (interactive review) →
   `do_reimport_apply` or `do_reimport_abort`.

## See also

- [CLI reference](cli.md) — the same surface, different transport.
- [Source format](source-format.md) — what the tools read and write.
- [Escape hatch tutorial](../tutorials/escape-hatch.md) — the
  re-import flow end-to-end.
