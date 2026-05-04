# claudesheets — design

## Purpose

Enable an economist to use Claude Code to build and maintain econometric
spreadsheet models with a developer-style workflow: TDD, progressive commits,
diffs, and review. The primary user is an economist who works in Microsoft
Excel; the toolkit is format-agnostic so it can grow to support other
spreadsheets later.

## Framing

Claude for Excel is an in-app assistant. It cannot be driven from Claude Code,
does not give the economist a TDD or git-style workflow, and inherits the
fundamental problems of `.xlsx` as a source format: zipped XML is not
diffable, mergeable, or grep-able.

claudesheets takes a different approach: **`.xlsx` is a build artifact, not
the source of truth.** The source of truth is a directory of text and CSV
files (with optional cached SQLite for fast queries). Claude Code edits the
source; a build step compiles it to `.xlsx` and a calc engine evaluates it.
Tests use Testsweet. Diffs are git diffs.

## Scope

### In scope (v1)

Excel features that round-trip with full fidelity:

**Tier 1 — Core**
- Cell values, formulas, data types, number formats
- Named ranges
- Multi-sheet workbooks
- Basic formatting (fonts, colors, borders, column widths)
- Data validation rules

**Tier 2 — Modeling essentials**
- Conditional formatting
- Cell comments / notes
- Frozen panes, print areas
- Excel "ListObject" tables (not "data tables")

### Out of scope (v1)

- Charts and pivot tables — economist regenerates these in Excel after build
- Macros / VBA (`.xlsm`) — explicit non-goal
- Excel "data tables" (what-if feature) — explicit non-goal
- ActiveX, form controls — explicit non-goal
- Multi-workbook projects and external references — single workbook per
  project; external refs in imported xlsx error by default

### Deferred to v2

- Watch-mode (`serve`) that rebuilds on save
- Source formatting / canonicalization (`format`)
- Per-sheet exporters (CSV, Markdown, JSON)
- Parquet for very large bulk data (>~100MB CSV equivalent)
- Multi-workbook projects, if real demand emerges

## Architecture

### Calc engine

Default: **LibreOffice Calc in headless mode** (`unoserver` /
`libreoffice --headless`). Free, scriptable, available on every platform,
covers Excel formula semantics well enough for econometric work.

The calc engine is a swappable plugin behind a small interface
(`evaluate(xlsx_path) -> calculated_values`). Future engines under the same
interface:

- `xlwings` driving a real Excel process (Windows/Mac local, full fidelity
  including macros if ever needed)
- Microsoft Graph workbook API (cloud `.xlsx`, server-side calc)
- Google Sheets API (if a Sheets path is added later)

Browser automation (Playwright) is explicitly **not** the integration
mechanism — the official APIs above are dramatically more reliable.

### Source format

A claudesheets project is a directory:

```
my-model/
├── claudesheets.toml          # project config: build settings, calc-engine choice
├── workbook.toml              # workbook-level: sheet order, named ranges, defined names
├── sheets/
│   ├── 01_assumptions.md      # one file per sheet — Markdown table for values
│   ├── 01_assumptions.yaml    # sidecar: formulas, formats, validation, conditional formatting
│   ├── 02_revenue.md
│   ├── 02_revenue.yaml
│   └── ...
├── data/
│   ├── cpi_series.csv         # bulk tabular data, text, diffable, committed
│   ├── panel_data.csv
│   └── _schema.sql            # optional column types / indexes for built SQLite
├── tests/
│   └── *.py                   # Testsweet tests
├── imports/                   # optional, opt-in: archived imported xlsx files
├── build/                     # gitignored: built .xlsx
└── .claudesheets/             # gitignored: built bulk.sqlite, calc cache
```

**Per-sheet `.md` file:** Markdown table holding values and references to
formula cells. The economist (and Claude) reads this to understand the sheet.

**Per-sheet `.yaml` sidecar:** structured data that doesn't fit a table —
formulas (keyed by both A1 address and named range; Excel forbids name/A1
collisions, so the namespaces don't clash), number formats, column widths,
data validation rules, conditional formatting rules, comments, frozen-pane
settings.

**Two files per sheet, not one combined file with frontmatter.** The `.md`
stays approachable for non-technical humans; the `.yaml` is where structure
lives.

**`workbook.toml`:** sheet order, named ranges with scope, workbook-level
styles.

**Bulk data:** CSV in `data/`. The build step loads CSVs into a cached
SQLite at `.claudesheets/bulk.sqlite` for fast SQL access during recalc. The
SQLite is gitignored — it's a build artifact, rebuilt deterministically from
CSV. Optional `_schema.sql` provides types and indexes. Bulk-data use is
opt-in per sheet (a sheet can declare `source: data/cpi_series.csv` and
draw values from it); small sheets stay fully in `.md`.

### Version control

| Versioned                                    | Ignored                                  |
|----------------------------------------------|------------------------------------------|
| `claudesheets.toml`, `workbook.toml`         | `build/` (built xlsx)                    |
| `sheets/*.md`, `sheets/*.yaml`               | `.claudesheets/` (built SQLite, caches)  |
| `data/*.csv`, `data/_schema.sql`             |                                          |
| `tests/*.py`                                 |                                          |
| `imports/*.xlsx` (opt-in via `--archive`)    |                                          |

## Workflows

### Initial onboarding (existing xlsx)

```
claudesheets import path/to/her-model.xlsx [--archive]
```

Reads the xlsx, populates `sheets/`, `workbook.toml`, and `data/`. With
`--archive`, copies the source xlsx into `imports/2026-04-25T1530.xlsx` for
traceability. External references in the imported xlsx error by default;
`--flatten` opts into replacing them with cached values.

### Greenfield project

```
claudesheets init my-model
```

Scaffolds an empty project. Claude builds it up from scratch in conversation.

### Edit / build / test loop

Claude edits source files directly with Read/Edit/Write — no fine-grained
mutator commands. The CLI surface stays coarse so source files are the API.

```
claudesheets build              # compile sources -> build/my-model.xlsx
claudesheets recalc             # run calc engine, cache calculated values
claudesheets test [-k pattern]  # run Testsweet tests
claudesheets snapshot [--update]  # golden-file regression of all calculated outputs
claudesheets diff [--vs xlsx:<path>]  # semantic diff
claudesheets check              # lint: dangling refs, missing names, schema mismatches
```

`recalc` is a separate command (not implicit in `build` or `test`) so a future
file watcher can trigger it on save and tests/builds can use cached values
based on file mtime.

### Testing model

**Primary path: Testsweet.** Tests are plain Python functions decorated
with `@test`, with no fixture-injection magic:

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

A small `claudesheets.testing` library exposes `Model.set/get/recalc` over
the chosen calc engine. [Testsweet](https://github.com/kaapstorm/testsweet)
does the heavy lifting; tests are explicit Python — no name-prefix
discovery, no fixture injection.

**Complementary path: golden-file snapshots.** `claudesheets snapshot`
computes all formula-cell outputs and compares to a checked-in golden file.
Catches regressions broadly. Snapshots are reviewed and accepted by running
`claudesheets snapshot --update`. The economist confirms calculated outputs
in the rebuilt workbook; Claude updates snapshots after she signs off.

### Escape hatch: economist edits the xlsx directly

The economist will sometimes edit `build/my-model.xlsx` directly between
Claude sessions. Documentation will recommend against this, but it must work.

**Detection.** `build` records the hash of the xlsx it produced. Any
subsequent claudesheets command checks the hash; if the file has changed,
it warns:

> `build/my-model.xlsx` has been modified externally. Run
> `claudesheets import build/my-model.xlsx` to review changes.

**Re-import flow.** Same `import` command as initial onboarding. When source
is non-empty, `import` becomes a review-first merge:

1. Re-imports the xlsx into a temp location.
2. Diffs against current source.
3. Prints a human-readable report ("B12 changed 0.04 → 0.05; 3 rows added to
   `revenue`").
4. In **interactive mode (default)**: presents three choices —
   - **Merge** (apply external changes on top of source; if source has
     edits since last build, Claude walks the user through any conflicts
     in plain language, with the option to back out)
   - **Overwrite** (discard source, accept xlsx as new source)
   - **Reject** (keep source unchanged, discard xlsx changes)
5. In **non-interactive mode (`-I`)**: prints the diff and exits, awaiting
   one of `import --apply` / `import --force` / `import --abort`.

We use "import" rather than "sync" because the operation is one-way: it
never modifies `build/my-model.xlsx`; it only updates source.

**Uncommitted source edits.** If `sheets/` has uncommitted changes,
`import` refuses unless `--force` is passed:

> You have uncommitted changes in `sheets/`. Commit or stash them before
> importing, or pass `--force` to discard.

**Conflict guidance.** Conflicts indicate the economist edited the xlsx
*and* Claude edited source since the last build. Documentation will
emphasize avoiding this. When it happens, Claude interprets the conflict
in plain language and walks the economist through resolution, always
offering the option to back out to Overwrite or Reject.

## Interfaces

### CLI

`claudesheets <command>`. Single binary, source of truth for behaviour.

| Command                  | Purpose                                              |
|--------------------------|------------------------------------------------------|
| `init [<path>]`          | Scaffold a fresh project                             |
| `import <xlsx>`          | Initial ingest, or review-first re-import            |
| `build [--out <path>]`   | Compile sources → `build/<name>.xlsx`                |
| `recalc`                 | Run calc engine, cache calculated values             |
| `test [-k <pattern>]`    | Run Testsweet tests                                  |
| `snapshot [--update]`    | Golden-file regression                               |
| `diff [--vs <ref>]`      | Semantic source-vs-source or source-vs-xlsx diff     |
| `check`                  | Lint: dangling refs, unused names, schema mismatches |

### MCP server

A thin wrapper around the CLI exposing the same commands as MCP tools with
typed inputs/outputs. The CLI is the source of truth; the MCP wrapper does
not duplicate logic. Anything Claude can do, the economist can reproduce in
a terminal.

## Out-of-scope clarifications

- **Browser automation.** Not used. Official APIs (LibreOffice headless,
  xlwings, Microsoft Graph, Sheets API) are the integration surface.
- **A new wire format for diffs.** Not needed. Source is text; `git diff`
  works. A semantic diff command sits on top, computing impact via the calc
  engine.
- **Replacing Excel.** The economist's deliverable is still an `.xlsx`.
  claudesheets is plumbing.

## Open questions for v2

- Watch-mode (`serve`) for auto-recalc on source change.
- Semantic-diff renderer that uses the calc engine to compute downstream
  impact ("growth_rate 4% → 5% changes Outputs!revenue_2027 by +24%").
- Multi-workbook projects, if real demand emerges.
- Sheets / cloud-Excel calc engines as alternatives to LibreOffice.
- `imports/` retention policy (keep all? prune to N? content-addressed?).
