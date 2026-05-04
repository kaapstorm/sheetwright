# claudesheets

A toolkit that lets [Claude Code](https://claude.com/claude-code) work with
spreadsheets the way it works with code: text-source files, git, TDD,
diffs, and review.

`.xlsx` is treated as a build artifact compiled from text sources.
[LibreOffice Calc](https://www.libreoffice.org/) in headless mode is the
default (swappable) calc engine.

See [`claude/specs/2026-04-25_claudesheets-design.md`](claude/specs/2026-04-25_claudesheets-design.md)
for the design.

## Status

Plan 5 complete: claudesheets is feature-complete for v1. The MCP
wrapper exposes every CLI command as an MCP tool with typed
inputs/outputs; `claudesheets mcp` launches the server on stdio.
Future work (v2): watch-mode, source canonicalisation, JSON diff
output, threaded-comments fidelity.

## Quick reference

```bash
claudesheets init my-model              # scaffold a fresh project
claudesheets import path/to/model.xlsx  # ingest an existing workbook (or re-import)
claudesheets build                      # compile sources -> build/<name>.xlsx
claudesheets recalc                     # run the calc engine, cache results
claudesheets test                       # run testsweet tests
claudesheets snapshot [--update]        # golden-file regression of outputs
claudesheets diff [--vs xlsx:<path>]    # semantic diff of source or vs an xlsx
claudesheets check                      # lint dangling refs, missing names, etc.
```

## Installation

claudesheets uses [`uv`](https://docs.astral.sh/uv/) for project management
and Python 3.11.

```bash
# Clone
git clone https://github.com/kaapstorm/claudesheets.git
cd claudesheets

# Install (creates .venv, installs runtime + dev deps)
uv sync

# Run the CLI
uv run claudesheets --help

# Run tests
uv run pytest
```

You will also need [LibreOffice](https://www.libreoffice.org/download/) on
your `PATH` for the default calc engine.

## Calc engine

The default engine is LibreOffice headless. `soffice` must be on
`$PATH` (Debian/Ubuntu: `apt install libreoffice`; macOS:
`brew install --cask libreoffice`).

The engine is selected per-project in `claudesheets.toml`:

```toml
[build]
calc_engine = "libreoffice"
```

The interface is documented in `src/claudesheets/calc/base.py`;
implement `CalcEngine.evaluate` to add a new backend.

## Testing your model

Tests use [testsweet](https://github.com/kaapstorm/testsweet) — plain
Python functions decorated with `@test`. Install claudesheets into
your project venv (not via `uv tool install`, which isolates
claudesheets from your project's dependencies):

```bash
uv add claudesheets
# or, if not using uv:
pip install claudesheets
```

Then write tests under `tests/`:

```python
import math

from testsweet import test

from claudesheets.testing import Model


@test
def revenue_grows_with_assumption():
    model = Model.open('.')
    model.set('Assumptions!growth_rate', 0.05)
    assert math.isclose(
        model.get('Outputs!revenue_2027'), 1_234_567, rel_tol=1e-6
    )
```

Run them with `claudesheets test` (in-process testsweet) or directly
with `python -m testsweet tests/`.

## Working with an externally-edited xlsx

If you (or your colleague) opens `build/<name>.xlsx` in Excel and
saves changes directly, claudesheets will notice on the next
command and warn:

    WARNING: build/my-model.xlsx has been modified externally.
    Run `claudesheets import build/my-model.xlsx` to review changes.

Re-importing into a populated source directory presents a diff and
asks how to proceed:

    claudesheets import build/my-model.xlsx
    # (m)erge / (o)verwrite / (r)eject

For non-interactive workflows (CI), use `-I` to stage the diff:

    claudesheets import build/my-model.xlsx -I
    # ... review the printed diff ...
    claudesheets import --apply   # accept
    claudesheets import --abort   # discard

If your `sheets/` directory has uncommitted git changes, re-import
refuses unless you pass `--force`.

## Diff and check

```bash
claudesheets diff                        # source vs build/<name>.xlsx
claudesheets diff --vs xlsx:other.xlsx   # source vs another xlsx
claudesheets diff --vs source:../other   # source vs another project's source
claudesheets check                       # lint dangling refs, missing names
```

`diff` exits 0 when there are no changes, 1 otherwise (CI-friendly).
`check` exits 0 when there are no issues, 1 otherwise.

## MCP server

claudesheets ships with an MCP (Model Context Protocol) server that
exposes every CLI command as a typed MCP tool, so Claude (or any
MCP client) can drive a project programmatically.

```bash
claudesheets mcp
```

Tools:

- `do_init(path)`
- `do_import_xlsx(xlsx, project, archive=False, flatten=False)`
- `do_build(project, out_path=None)`
- `do_recalc(project, force=False)`
- `do_snapshot(project, update=False)` — returns `{ok, message, has_diffs}`
- `do_test(project, targets=[])` — returns `{passed, output}`
- `do_diff(project, vs=None)` — returns `{is_empty, rendered, structured}`
- `do_check(project)` — returns `{issues: [{kind, detail, location}, ...]}`
- `do_reimport_stage(xlsx, project, flatten=False, force=False)` —
  returns the diff and saves a session.
- `do_reimport_apply(project, archive=False)` — completes a staged session.
- `do_reimport_abort(project)` — discards a staged session.

The server runs over stdio. Most MCP clients launch
`claudesheets mcp` as a subprocess and route MCP traffic over
stdin/stdout.

## License

[GPL-3.0-or-later](LICENSE).
