# sheetwright

A toolkit that lets [Claude Code](https://claude.com/claude-code) work with
spreadsheets the way it works with code: text-source files, git, TDD,
diffs, and review.

`.xlsx` is treated as a build artifact compiled from text sources. You
write Markdown tables and YAML sidecars; sheetwright produces the
workbook and runs [LibreOffice Calc](https://www.libreoffice.org/) in
headless mode (the default, swappable calc engine) to evaluate every
formula. From there you assert on outputs with ordinary Python tests,
diff source against any other workbook, and snapshot calculated values
for regression checks.

sheetwright also ships an MCP (Model Context Protocol) server, so
Claude Code (or any MCP client) can drive a project end-to-end through
typed tools that mirror the CLI.

## Status

v1.0.0. See [`docs/`](docs/) for the manual.

## Quick reference

```bash
sheetwright init my-model              # scaffold a fresh project
sheetwright import path/to/model.xlsx  # ingest an existing workbook (or re-import)
sheetwright build                      # compile sources -> build/<name>.xlsx
sheetwright recalc                     # run the calc engine, cache results
sheetwright test                       # run testsweet tests
sheetwright snapshot [--update]        # golden-file regression of outputs
sheetwright diff [--vs xlsx:<path>]    # semantic diff of source or vs an xlsx
sheetwright check                      # lint dangling refs, missing names, etc.
sheetwright mcp                        # run the MCP server on stdio
```

## Installation

sheetwright uses [`uv`](https://docs.astral.sh/uv/) for project
management and Python 3.11.

```bash
git clone https://github.com/kaapstorm/sheetwright.git
cd sheetwright
uv sync
uv run sheetwright --help
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
  `sheetwright.testing.Model` API and testsweet patterns.
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
[`claude/specs/2026-04-25_sheetwright-design.md`](claude/specs/2026-04-25_sheetwright-design.md).

## License

[GPL-3.0-or-later](LICENSE).
