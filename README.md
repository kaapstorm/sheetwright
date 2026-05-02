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

Plan 3 complete: Tier 2 features (conditional formatting, comments,
frozen panes, print areas, ListObject tables) round-trip with full
fidelity through `import → build`. The escape-hatch re-import flow
and `diff`/`check` commands are coming in Plan 4.

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

## License

[GPL-3.0-or-later](LICENSE).
