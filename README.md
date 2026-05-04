# claudesheets

A toolkit that lets [Claude Code](https://claude.com/claude-code) work with
spreadsheets the way it works with code: text-source files, git, TDD,
diffs, and review.

`.xlsx` is treated as a build artefact compiled from text sources. You
write Markdown tables and YAML sidecars; claudesheets produces the
workbook and runs [LibreOffice Calc](https://www.libreoffice.org/) in
headless mode (the default, swappable calc engine) to evaluate every
formula. From there you assert on outputs with ordinary Python tests,
diff source against any other workbook, and snapshot calculated values
for regression checks.

claudesheets also ships an MCP (Model Context Protocol) server, so
Claude Code (or any MCP client) can drive a project end-to-end through
typed tools that mirror the CLI.

## Status

v1.0.0. See [`docs/`](docs/) for the manual.

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
claudesheets mcp                        # run the MCP server on stdio
```

## Installation

claudesheets uses [`uv`](https://docs.astral.sh/uv/) for project
management and Python 3.11.

```bash
git clone https://github.com/kaapstorm/claudesheets.git
cd claudesheets
uv sync
uv run claudesheets --help
```

LibreOffice must be on `$PATH` for the default calc engine
(Debian/Ubuntu: `apt install libreoffice`; macOS:
`brew install --cask libreoffice`).

## Documentation

The manual lives under [`docs/`](docs/):

- [Getting started](docs/getting-started.md) — install through first
  test in about ten minutes.
- [CLI reference](docs/reference/cli.md) — every command and flag.
- [Source format](docs/reference/source-format.md) — on-disk layout
  and the Markdown / YAML / TOML semantics.
- [Testing reference](docs/reference/testing.md) — the
  `claudesheets.testing.Model` API and testsweet patterns.
- [MCP reference](docs/reference/mcp.md) — typed tools, return
  shapes, error codes.
- [Calc engine reference](docs/reference/calc-engine.md) — the
  plugin interface and the LibreOffice backend.
- Tutorials: [greenfield](docs/tutorials/greenfield-project.md),
  [importing](docs/tutorials/importing-existing.md),
  [TDD](docs/tutorials/tdd-workflow.md),
  [snapshots](docs/tutorials/snapshots.md),
  [escape hatch](docs/tutorials/escape-hatch.md).

## Design

The design rationale is in
[`claude/specs/2026-04-25_claudesheets-design.md`](claude/specs/2026-04-25_claudesheets-design.md).

## License

[GPL-3.0-or-later](LICENSE).
