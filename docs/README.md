# sheetwright manual

sheetwright treats `.xlsx` as a build artifact compiled from text
sources. You write Markdown tables and YAML sidecars; sheetwright
produces a workbook, runs LibreOffice (or another calc engine) to
evaluate every formula, and lets you assert on the results from
ordinary Python tests. The aim: spreadsheets that you can review in a
pull request, regression-test in CI, and edit alongside Claude Code.

## Getting started

- [Getting started](getting-started.md) — install, scaffold a
  project, edit a sheet, build, and run your first test in about ten
  minutes.

## Reference

- [CLI](reference/cli.md) — every command and flag.
- [Source format](reference/source-format.md) — on-disk layout, the
  Markdown table format, the YAML sidecar, and the two manifests.
- [Testing](reference/testing.md) — `sheetwright.testing.Model` and
  testsweet patterns.
- [MCP server](reference/mcp.md) — typed tools, return shapes, error
  codes, and how to wire it into a client.
- [Calc engine](reference/calc-engine.md) — the `CalcEngine`
  interface, how the cache works, and how to add a new backend.

## Tutorials

- [Greenfield project](tutorials/greenfield-project.md) — build a new
  model from scratch.
- [Importing an existing workbook](tutorials/importing-existing.md) —
  onboard an existing `.xlsx`.
- [TDD workflow](tutorials/tdd-workflow.md) — test-driven model
  development.
- [Snapshots](tutorials/snapshots.md) — golden-file regression for
  calculated outputs.
- [Escape hatch](tutorials/escape-hatch.md) — handle direct edits made
  to the built `.xlsx`.

## Feature matrix

| Tier | What's covered |
| --- | --- |
| Tier 1 | Cell values, formulas, named ranges, sheets, column widths, number formats, fonts, fills, borders, comments, data validation, list tables. |
| Tier 2 | Conditional formatting (cell-is, formula, color scale, data bar, icon set), frozen panes, print area. |
| Build  | Deterministic xlsx writer, LibreOffice-headless evaluator, content-addressed calc cache. |
| MCP    | Twelve typed tools that mirror the CLI; ships in the same package. |

## Why

The design rationale lives in
`claude/specs/2026-04-25_sheetwright-design.md`. The short version:
spreadsheets are software, but they're missing diff, review, and
tests. sheetwright fills in the missing tooling without forcing you
off Excel.
