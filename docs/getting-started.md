# Getting started

Ten minutes from `pip install` to a passing test.

## Prerequisites

- Python 3.11.
- [LibreOffice](https://www.libreoffice.org/download/) on `$PATH`
  (provides `soffice`, the default calc engine).
- Git (recommended; sheetwright generates a `.gitignore` and the
  re-import flow checks for uncommitted source changes).

On Debian/Ubuntu:

```bash
sudo apt install libreoffice python3.11
```

On macOS:

```bash
brew install --cask libreoffice
brew install python@3.11
```

On Windows (PowerShell, using
[WinGet](https://learn.microsoft.com/en-us/windows/package-manager/winget/)):

```powershell
winget install --id TheDocumentFoundation.LibreOffice
winget install --id Python.Python.3.11
winget install --id Git.Git
```

New to PowerShell or `winget`? Start with the [Windows setup
primer](tutorials/windows-setup.md), which covers `winget`,
LibreOffice, Sourcetree, and the PowerShell commands you'll need.

## Install

sheetwright is built on [`uv`](https://docs.astral.sh/uv/). The
quickest path for a fresh project:

```bash
uv init my-model
cd my-model
uv add sheetwright
```

If you prefer pip:

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install sheetwright
```

On Windows:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install sheetwright
```

If PowerShell blocks `Activate.ps1` with an execution-policy error,
allow signed local scripts for your user once with:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Confirm:

```bash
$ uv run sheetwright --version
sheetwright, version 0.1.0
```

## Scaffold a project

```bash
$ uv run sheetwright init .
Initialised sheetwright project at /home/you/my-model
```

This creates:

```
my-model/
├── sheetwright.toml      # project config (name, calc engine)
├── workbook.toml          # workbook manifest (sheets, named ranges)
├── sheets/                # one .md (+ optional .yaml) per sheet
├── data/                  # bulk CSVs (loaded into a sqlite cache on build)
├── tests/
│   └── __init__.py
└── .gitignore             # ignores build/ and .sheetwright/
```

See [source format](reference/source-format.md) for the full layout.

If you already have an `.xlsx`, use `sheetwright import path/to.xlsx`
instead — see [importing an existing
workbook](tutorials/importing-existing.md).

## Add a sheet

Edit `workbook.toml` to register two sheets:

```toml
[workbook]
name = "my-model"
sheets = ["Inputs", "Outputs"]

[[named_ranges]]
name = "growth_rate"
scope = "workbook"
ref = "Inputs!$B$1"
```

Create `sheets/01_inputs.md`:

```markdown
| (cell) | A             | B    |
| ---    | ---           | ---  |
| 1      | growth_rate   | 0.04 |
| 2      | base_revenue  | 1000000 |
```

Create `sheets/02_outputs.md`:

```markdown
| (cell) | A           | B                         |
| ---    | ---         | ---                       |
| 1      | revenue_y1  | =Inputs!B2*(1+growth_rate) |
```

The first column header is the literal string `(cell)`; subsequent
columns are Excel-style column letters. Formulas start with `=`. See
[source format](reference/source-format.md) for the full grammar.

## Build

```bash
$ uv run sheetwright build
Built /home/you/my-model/build/my-model.xlsx
```

Open it in Excel or LibreOffice to confirm the formulas wrote
through. The xlsx is reproducible: rebuilding from the same source
produces a byte-identical file.

## Recalc

```bash
$ uv run sheetwright recalc
recalculated: /home/you/my-model/.sheetwright/calc/<sha>.json
```

`recalc` runs LibreOffice headless against the built xlsx and caches
the result keyed by the xlsx's SHA-256. Subsequent runs hit the
cache:

```bash
$ uv run sheetwright recalc
cache hit: 4f1a9c2e8b7d
```

## Write a test

Create `tests/test_revenue.py`:

```python
import math

from testsweet import test

from sheetwright.testing import Model


@test
def revenue_grows_with_assumption():
    model = Model.open('.')
    model.set('growth_rate', 0.05)
    expected = 1_000_000 * 1.05
    assert math.isclose(model.get('Outputs!B1'), expected, rel_tol=1e-9)
```

`Model.open` reads source from disk; `set` mutates the in-memory
workbook; `get` triggers a recalc on demand and returns the
calculated value. See [testing reference](reference/testing.md) for
the full API.

## Run the test

```bash
$ uv run sheetwright test
tests/test_revenue.py::revenue_grows_with_assumption ... ok
```

That's the loop. Edit source, build, recalc, test — every step is
text-in, text-out, and every artefact under `build/` and
`.sheetwright/` can be regenerated from `sheets/` and the manifests.

## Where to next?

- [Greenfield tutorial](tutorials/greenfield-project.md) walks
  through a fuller model from scratch (formats, named ranges,
  validation).
- [TDD workflow](tutorials/tdd-workflow.md) shows how to write the
  test before the formula.
- [Snapshots](tutorials/snapshots.md) catches regressions across
  every formula cell, not just the ones you remembered to assert on.
- [CLI reference](reference/cli.md) lists every command.
- [MCP reference](reference/mcp.md) covers the Claude-facing server.
