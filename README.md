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

Plan 1 complete: Tier 1 round-trip foundation. `init`, `import`, `build`
work end-to-end for values, formulas, named ranges, basic formatting,
number formats, and data validation. Calc engine, tests, snapshot,
diff, and the escape-hatch import flow are coming in Plans 2–4.

## Quick reference

```bash
claudesheets init my-model              # scaffold a fresh project
claudesheets import path/to/model.xlsx  # ingest an existing workbook (or re-import)
claudesheets build                      # compile sources -> build/<name>.xlsx
claudesheets recalc                     # run the calc engine, cache results
claudesheets test                       # run pytest tests
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

## License

[GPL-3.0-or-later](LICENSE).
