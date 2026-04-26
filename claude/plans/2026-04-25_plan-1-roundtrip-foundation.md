# Plan 1 — Round-trip foundation

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the foundation that lets a Tier 1 `.xlsx` round-trip losslessly through claudesheets' text source format. End state: `claudesheets import path/to/model.xlsx && claudesheets build` produces an `.xlsx` semantically equivalent to the input.

**Architecture:** Three layers — an in-memory `Workbook` model; readers/writers between `Workbook` and `.xlsx` (using `openpyxl`); readers/writers between `Workbook` and the text source directory (Markdown for value tables, YAML for formats/validation, TOML for the workbook manifest). Bulk CSV data is loaded into a gitignored cached SQLite at build time. CLI commands (`init`, `import`, `build`) are thin wrappers around these layers.

**Tech Stack:** Python 3.11, `uv`, `pytest` + `pytest-unmagic`, `openpyxl` (xlsx I/O), `ruamel.yaml` (sidecar YAML), `tomli-w` (TOML writer; reading uses stdlib `tomllib`), `click` (CLI).

**Scope reminder.** Tier 1 features only:
- Cell values (numbers, strings, booleans, dates, blanks)
- Formulas
- Number formats
- Named ranges (workbook + sheet scope)
- Multi-sheet workbooks
- Basic formatting (fonts, fills, borders, column widths)
- Data validation rules

External references in imports raise an error (no `--flatten` in this plan; deferred to Plan 4). Charts/pivots/macros explicitly skipped — `import` warns and drops them. Escape hatch and `recalc`/`test`/`diff`/`check` commands are Plans 2–4.

**Source format (locked by the design spec, restated here for reference).**

```
my-model/
├── claudesheets.toml          # project config (calc engine, build name)
├── workbook.toml              # workbook manifest (sheet order, named ranges)
├── sheets/
│   ├── 01_assumptions.md      # Markdown table; values or formula text (=...)
│   ├── 01_assumptions.yaml    # column widths, formats, validation
│   └── ...
├── data/
│   ├── *.csv                  # bulk tabular data (optional)
│   └── _schema.sql            # optional column types
├── tests/                     # (Plan 2)
├── imports/                   # opt-in archive of imported xlsx files
├── build/                     # gitignored: built xlsx
└── .claudesheets/             # gitignored: built bulk.sqlite, caches
```

**For Plan 1 only:** formulas live in the `.md` table as literal text (starting with `=`). The `.yaml` does *not* hold formulas yet. Plan 2, when calculated values exist, will move formulas to `.yaml` and put calc values in `.md`. This split is deliberate — keeps Plan 1 simple, and the test we'll write for round-trip equivalence will continue to pass after Plan 2's reorganization.

---

## File structure

```
src/claudesheets/
├── __init__.py                 (exists)
├── cli.py                      Click group, dispatches to commands/*
├── exceptions.py               ClaudesheetsError hierarchy
├── project.py                  Project class: paths, config loading
├── commands/
│   ├── __init__.py
│   ├── init_cmd.py             `claudesheets init`
│   ├── import_cmd.py           `claudesheets import` (initial-only in Plan 1)
│   └── build_cmd.py            `claudesheets build`
├── model/
│   ├── __init__.py
│   ├── workbook.py             Workbook, Sheet, NamedRange dataclasses
│   ├── cell.py                 Cell dataclass (value, formula, format ref)
│   ├── format.py               CellFormat, Font, Fill, Border, ColumnWidth
│   └── validation.py           DataValidation dataclass
├── xlsx/
│   ├── __init__.py
│   ├── reader.py               xlsx → Workbook
│   └── writer.py               Workbook → xlsx
├── source/
│   ├── __init__.py
│   ├── reader.py               source dir → Workbook
│   ├── writer.py               Workbook → source dir
│   └── markdown.py             Markdown-table parse/serialize
├── bulk.py                     data/*.csv → .claudesheets/bulk.sqlite
└── config.py                   claudesheets.toml + workbook.toml schemas

tests/
├── __init__.py
├── conftest.py                 (intentionally minimal; fixtures imported, not auto-discovered)
├── fixtures/
│   ├── __init__.py
│   ├── workbooks.py            Programmatic Workbook builders for tests
│   └── files/                  small .xlsx files used as test inputs
├── test_xlsx_roundtrip.py
├── test_source_markdown.py
├── test_source_roundtrip.py
├── test_bulk.py
├── test_init.py
├── test_import.py
├── test_build.py
└── test_end_to_end.py
```

---

## Conventions for every task

- **Use `pytest-unmagic`.** All fixtures defined with `@fixture`; applied with `@use(...)` or shorthand. Never bare `pytest.fixture`. See CLAUDE.md.
- **Each task ends with a commit.** Commit message format: `<area>: <imperative summary>` (e.g. `xlsx: read multi-sheet workbooks`). No co-author trailer needed within plan execution.
- **Run from repo root.** All commands assume `cd /srv/share/src/kaapstorm/claudesheets` (or equivalent worktree).
- **Use `uv run`** to invoke `pytest` and the `claudesheets` CLI. Never `pip install` or `python -m pytest` directly.

---

## Task 0: Add runtime dependencies

**Files:**
- Modify: `pyproject.toml`

- [ ] **Step 1: Add runtime dependencies via `uv`**

```bash
uv add openpyxl 'ruamel.yaml' tomli-w click
```

Expected: `pyproject.toml`'s `dependencies` list updated; `uv.lock` regenerated; `.venv` updated.

- [ ] **Step 2: Verify CLI still imports**

```bash
uv run python -c "import claudesheets; import openpyxl; import ruamel.yaml; import tomli_w; import click; print('ok')"
```

Expected: `ok`.

- [ ] **Step 3: Commit**

```bash
git add pyproject.toml uv.lock
git commit -m "deps: add openpyxl, ruamel.yaml, tomli-w, click"
```

---

## Task 1: CLI skeleton with `click`

**Files:**
- Modify: `src/claudesheets/cli.py`
- Create: `src/claudesheets/exceptions.py`
- Create: `src/claudesheets/commands/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Create: `tests/test_cli.py`

- [ ] **Step 1: Write the failing test**

Create `tests/__init__.py` (empty) and `tests/conftest.py` (empty — pytest-unmagic does not need it for autodiscovery, but pytest expects the file to exist for test rootdir resolution in some configs; an empty file is fine).

Create `tests/test_cli.py`:

```python
from click.testing import CliRunner

from claudesheets.cli import main


def test_help_lists_subcommands():
    runner = CliRunner()
    result = runner.invoke(main, ["--help"])
    assert result.exit_code == 0
    for cmd in ("init", "import", "build"):
        assert cmd in result.output


def test_unknown_subcommand_errors():
    runner = CliRunner()
    result = runner.invoke(main, ["bogus"])
    assert result.exit_code != 0
```

- [ ] **Step 2: Run tests; confirm they fail**

```bash
uv run pytest tests/test_cli.py -v
```

Expected: both tests fail (current `main` raises `NotImplementedError`).

- [ ] **Step 3: Create the exception module**

`src/claudesheets/exceptions.py`:

```python
"""Exception hierarchy for claudesheets."""


class ClaudesheetsError(Exception):
    """Base class for all claudesheets errors."""


class ProjectError(ClaudesheetsError):
    """A project-level error (missing config, malformed structure)."""


class ImportError_(ClaudesheetsError):
    """Errors during xlsx import."""


class BuildError(ClaudesheetsError):
    """Errors during build."""
```

- [ ] **Step 4: Implement the CLI skeleton**

`src/claudesheets/commands/__init__.py`: empty.

`src/claudesheets/cli.py`:

```python
"""claudesheets CLI entry point."""
from __future__ import annotations

import sys

import click


@click.group(context_settings={"help_option_names": ["-h", "--help"]})
@click.version_option(package_name="claudesheets")
def main() -> None:
    """Work with spreadsheets from Claude Code."""


@main.command("init")
@click.argument("path", type=click.Path(file_okay=False), default=".")
def init_cmd(path: str) -> None:
    """Scaffold an empty claudesheets project."""
    from claudesheets.commands.init_cmd import run

    run(path)


@main.command("import")
@click.argument("xlsx", type=click.Path(exists=True, dir_okay=False))
@click.option("--archive", is_flag=True, help="Copy the xlsx into imports/.")
@click.option("--project", "project_path", type=click.Path(file_okay=False),
              default=".", help="Path to the claudesheets project.")
def import_cmd(xlsx: str, archive: bool, project_path: str) -> None:
    """Read an .xlsx file into source form."""
    from claudesheets.commands.import_cmd import run

    run(xlsx_path=xlsx, project_path=project_path, archive=archive)


@main.command("build")
@click.option("--out", "out_path", type=click.Path(dir_okay=False), default=None,
              help="Output path. Defaults to build/<workbook-name>.xlsx.")
@click.option("--project", "project_path", type=click.Path(file_okay=False),
              default=".", help="Path to the claudesheets project.")
def build_cmd(out_path: str | None, project_path: str) -> None:
    """Compile source files into an .xlsx."""
    from claudesheets.commands.build_cmd import run

    run(project_path=project_path, out_path=out_path)


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
```

The `run(...)` functions in `init_cmd.py`, `import_cmd.py`, `build_cmd.py` will be added in later tasks. For now, create stubs:

`src/claudesheets/commands/init_cmd.py`:

```python
def run(path: str) -> None:
    raise NotImplementedError
```

`src/claudesheets/commands/import_cmd.py`:

```python
def run(*, xlsx_path: str, project_path: str, archive: bool) -> None:
    raise NotImplementedError
```

`src/claudesheets/commands/build_cmd.py`:

```python
def run(*, project_path: str, out_path: str | None) -> None:
    raise NotImplementedError
```

- [ ] **Step 5: Run tests; confirm they pass**

```bash
uv run pytest tests/test_cli.py -v
```

Expected: both pass.

- [ ] **Step 6: Commit**

```bash
git add src/claudesheets/ tests/
git commit -m "cli: scaffold click subcommand dispatch"
```

---

## Task 2: `claudesheets init`

**Files:**
- Modify: `src/claudesheets/commands/init_cmd.py`
- Create: `tests/test_init.py`

- [ ] **Step 1: Write the failing test**

`tests/test_init.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main


@fixture
def tmp_project(tmp_path):
    yield tmp_path / "my-model"


@use(tmp_project)
def test_init_creates_project_skeleton():
    project = tmp_project()
    runner = CliRunner()
    result = runner.invoke(main, ["init", str(project)])
    assert result.exit_code == 0, result.output

    assert (project / "claudesheets.toml").is_file()
    assert (project / "workbook.toml").is_file()
    assert (project / "sheets").is_dir()
    assert (project / "data").is_dir()
    assert (project / ".gitignore").is_file()

    gitignore = (project / ".gitignore").read_text()
    assert "build/" in gitignore
    assert ".claudesheets/" in gitignore


@use(tmp_project)
def test_init_refuses_non_empty_directory():
    project = tmp_project()
    project.mkdir()
    (project / "stuff.txt").write_text("hi")

    runner = CliRunner()
    result = runner.invoke(main, ["init", str(project)])
    assert result.exit_code != 0
    assert "not empty" in result.output.lower()
```

- [ ] **Step 2: Run tests; confirm they fail**

```bash
uv run pytest tests/test_init.py -v
```

Expected: both tests fail (`NotImplementedError`).

- [ ] **Step 3: Implement `init`**

`src/claudesheets/commands/init_cmd.py`:

```python
"""Implementation of `claudesheets init`."""
from __future__ import annotations

from pathlib import Path

import click

DEFAULT_CLAUDESHEETS_TOML = """\
[project]
name = "my-model"

[build]
calc_engine = "libreoffice"
"""

DEFAULT_WORKBOOK_TOML = """\
[workbook]
name = "my-model"

# Sheets are listed in workbook order. Each entry must match a file in sheets/
# (without the .md/.yaml extension).
sheets = []

# Workbook-scoped named ranges:
# [[named_ranges]]
# name = "growth_rate"
# scope = "workbook"
# ref = "Assumptions!B5"
"""

DEFAULT_GITIGNORE = """\
build/
.claudesheets/
"""


def run(path: str) -> None:
    project = Path(path).resolve()
    if project.exists() and any(project.iterdir()):
        raise click.ClickException(f"{project} is not empty.")

    project.mkdir(parents=True, exist_ok=True)
    (project / "sheets").mkdir()
    (project / "data").mkdir()
    (project / "claudesheets.toml").write_text(DEFAULT_CLAUDESHEETS_TOML)
    (project / "workbook.toml").write_text(DEFAULT_WORKBOOK_TOML)
    (project / ".gitignore").write_text(DEFAULT_GITIGNORE)

    click.echo(f"Initialised claudesheets project at {project}")
```

- [ ] **Step 4: Run tests; confirm they pass**

```bash
uv run pytest tests/test_init.py -v
```

Expected: both pass.

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/commands/init_cmd.py tests/test_init.py
git commit -m "init: scaffold an empty project directory"
```

---

## Task 3: In-memory workbook model

**Files:**
- Create: `src/claudesheets/model/__init__.py`
- Create: `src/claudesheets/model/workbook.py`
- Create: `src/claudesheets/model/cell.py`
- Create: `tests/test_model.py`

- [ ] **Step 1: Write the failing test**

`tests/test_model.py`:

```python
from datetime import datetime

from claudesheets.model.cell import Cell
from claudesheets.model.workbook import NamedRange, Sheet, Workbook


def test_workbook_is_empty_by_default():
    wb = Workbook(name="m")
    assert wb.sheets == []
    assert wb.named_ranges == []


def test_sheet_get_set_round_trip():
    sh = Sheet(name="S1")
    sh.set("A1", Cell(value=10))
    sh.set("B2", Cell(formula="=A1*2"))
    assert sh.get("A1").value == 10
    assert sh.get("B2").formula == "=A1*2"


def test_sheet_get_unset_returns_blank_cell():
    sh = Sheet(name="S1")
    assert sh.get("Z99") == Cell()


def test_cell_supports_basic_types():
    Cell(value=1)
    Cell(value=1.5)
    Cell(value="x")
    Cell(value=True)
    Cell(value=datetime(2026, 1, 1))
    Cell(value=None)


def test_cell_cannot_have_value_and_formula_together():
    import pytest
    with pytest.raises(ValueError):
        Cell(value=1, formula="=A1")


def test_workbook_lookup_sheet_by_name():
    wb = Workbook(name="m")
    s = Sheet(name="Inputs")
    wb.sheets.append(s)
    assert wb.sheet("Inputs") is s


def test_named_range_workbook_scope():
    nr = NamedRange(name="growth", scope="workbook", ref="Inputs!B5")
    assert nr.scope == "workbook"


def test_named_range_sheet_scope_requires_sheet_name():
    import pytest
    with pytest.raises(ValueError):
        NamedRange(name="local", scope="sheet", ref="B5")  # missing sheet
    nr = NamedRange(name="local", scope="sheet", sheet="Inputs", ref="B5")
    assert nr.sheet == "Inputs"
```

- [ ] **Step 2: Run; confirm failures**

```bash
uv run pytest tests/test_model.py -v
```

Expected: import errors / failures (modules not yet created).

- [ ] **Step 3: Create the model**

`src/claudesheets/model/__init__.py`: empty.

`src/claudesheets/model/cell.py`:

```python
"""In-memory representation of a single spreadsheet cell."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, Union

CellValue = Union[None, bool, int, float, str, datetime]


@dataclass(frozen=True)
class Cell:
    """A single cell.

    Either `value` is set (a literal cell) or `formula` is set (a formula
    cell), or both are None (a blank cell). They cannot both be set.
    The `format_id` is a reference into the sheet's format table; it
    is set by readers and resolved by writers.
    """
    value: CellValue = None
    formula: Optional[str] = None
    format_id: Optional[str] = None

    def __post_init__(self) -> None:
        if self.value is not None and self.formula is not None:
            raise ValueError("Cell cannot have both value and formula.")

    @property
    def is_blank(self) -> bool:
        return self.value is None and self.formula is None and self.format_id is None
```

`src/claudesheets/model/workbook.py`:

```python
"""In-memory workbook, sheets, and named ranges."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from claudesheets.model.cell import Cell


@dataclass
class NamedRange:
    name: str
    ref: str
    scope: str = "workbook"  # "workbook" or "sheet"
    sheet: Optional[str] = None  # required when scope == "sheet"

    def __post_init__(self) -> None:
        if self.scope not in ("workbook", "sheet"):
            raise ValueError(f"Invalid scope: {self.scope!r}")
        if self.scope == "sheet" and not self.sheet:
            raise ValueError("Sheet-scoped named range requires `sheet`.")


@dataclass
class Sheet:
    name: str
    cells: Dict[str, Cell] = field(default_factory=dict)
    column_widths: Dict[str, float] = field(default_factory=dict)
    frozen_panes: Optional[str] = None  # e.g. "B2"; reserved for Plan 3

    def get(self, address: str) -> Cell:
        return self.cells.get(address, Cell())

    def set(self, address: str, cell: Cell) -> None:
        if cell.is_blank:
            self.cells.pop(address, None)
        else:
            self.cells[address] = cell

    def addresses(self) -> List[str]:
        return list(self.cells.keys())


@dataclass
class Workbook:
    name: str
    sheets: List[Sheet] = field(default_factory=list)
    named_ranges: List[NamedRange] = field(default_factory=list)

    def sheet(self, name: str) -> Sheet:
        for s in self.sheets:
            if s.name == name:
                return s
        raise KeyError(name)
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_model.py -v
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/model tests/test_model.py
git commit -m "model: add Workbook, Sheet, Cell, NamedRange"
```

---

## Task 4: xlsx → Workbook reader (values, formulas, multi-sheet)

**Files:**
- Create: `src/claudesheets/xlsx/__init__.py`
- Create: `src/claudesheets/xlsx/reader.py`
- Create: `tests/fixtures/__init__.py`
- Create: `tests/fixtures/workbooks.py`
- Create: `tests/test_xlsx_reader.py`

- [ ] **Step 1: Add a fixture builder for programmatic test workbooks**

`tests/fixtures/__init__.py`: empty.

`tests/fixtures/workbooks.py`:

```python
"""Helpers that build small openpyxl workbooks for round-trip tests.

These are not pytest-unmagic fixtures — they are plain helpers callable
from tests. We keep them in one place so test setup stays consistent.
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName


def write_simple_xlsx(path: Path) -> None:
    """A two-sheet workbook with values, a formula, and a named range."""
    wb = openpyxl.Workbook()
    s1 = wb.active
    s1.title = "Inputs"
    s1["A1"] = "growth_rate"
    s1["B1"] = 0.04
    s1["A2"] = "base_revenue"
    s1["B2"] = 1_000_000

    s2 = wb.create_sheet("Outputs")
    s2["A1"] = "revenue_2027"
    s2["B1"] = "=Inputs!B2 * (1 + Inputs!B1)"

    wb.defined_names["growth_rate"] = DefinedName(
        name="growth_rate", attr_text="Inputs!$B$1"
    )

    wb.save(path)
```

- [ ] **Step 2: Write the failing test**

`tests/test_xlsx_reader.py`:

```python
from pathlib import Path

from unmagic import fixture, use

from claudesheets.xlsx.reader import read_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def simple_xlsx(tmp_path):
    path = tmp_path / "simple.xlsx"
    write_simple_xlsx(path)
    yield path


@use(simple_xlsx)
def test_read_returns_workbook_with_two_sheets():
    wb = read_xlsx(simple_xlsx())
    assert wb.name == "simple"
    assert [s.name for s in wb.sheets] == ["Inputs", "Outputs"]


@use(simple_xlsx)
def test_read_preserves_values():
    wb = read_xlsx(simple_xlsx())
    inputs = wb.sheet("Inputs")
    assert inputs.get("A1").value == "growth_rate"
    assert inputs.get("B1").value == 0.04
    assert inputs.get("B2").value == 1_000_000


@use(simple_xlsx)
def test_read_preserves_formula():
    wb = read_xlsx(simple_xlsx())
    outputs = wb.sheet("Outputs")
    cell = outputs.get("B1")
    assert cell.value is None
    assert cell.formula == "=Inputs!B2 * (1 + Inputs!B1)"


@use(simple_xlsx)
def test_read_preserves_named_range():
    wb = read_xlsx(simple_xlsx())
    names = {nr.name: nr for nr in wb.named_ranges}
    assert "growth_rate" in names
    assert names["growth_rate"].ref == "Inputs!$B$1"
    assert names["growth_rate"].scope == "workbook"
```

- [ ] **Step 3: Run; confirm failures**

```bash
uv run pytest tests/test_xlsx_reader.py -v
```

Expected: import errors (reader not yet implemented).

- [ ] **Step 4: Implement the reader**

`src/claudesheets/xlsx/__init__.py`: empty.

`src/claudesheets/xlsx/reader.py`:

```python
"""Read an .xlsx file into a Workbook."""
from __future__ import annotations

from pathlib import Path

import openpyxl

from claudesheets.model.cell import Cell
from claudesheets.model.workbook import NamedRange, Sheet, Workbook


def read_xlsx(path: Path) -> Workbook:
    path = Path(path)
    src = openpyxl.load_workbook(path, data_only=False)

    wb = Workbook(name=path.stem)

    for ws in src.worksheets:
        sheet = Sheet(name=ws.title)
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if isinstance(c.value, str) and c.value.startswith("="):
                    sheet.set(c.coordinate, Cell(formula=c.value))
                else:
                    sheet.set(c.coordinate, Cell(value=c.value))
        wb.sheets.append(sheet)

    for name, defn in src.defined_names.items():
        # openpyxl exposes workbook-scope names via wb.defined_names and
        # sheet-scope names via ws.defined_names. We handle both.
        wb.named_ranges.append(
            NamedRange(name=name, ref=defn.attr_text, scope="workbook")
        )

    for ws in src.worksheets:
        for name, defn in ws.defined_names.items():
            wb.named_ranges.append(
                NamedRange(name=name, ref=defn.attr_text, scope="sheet", sheet=ws.title)
            )

    return wb
```

- [ ] **Step 5: Run; confirm pass**

```bash
uv run pytest tests/test_xlsx_reader.py -v
```

Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add src/claudesheets/xlsx tests/fixtures tests/test_xlsx_reader.py
git commit -m "xlsx: read values, formulas, sheets, and named ranges"
```

---

## Task 5: xlsx writer (values, formulas, multi-sheet)

**Files:**
- Create: `src/claudesheets/xlsx/writer.py`
- Create: `tests/test_xlsx_roundtrip.py`

- [ ] **Step 1: Write the failing round-trip test**

`tests/test_xlsx_roundtrip.py`:

```python
from pathlib import Path

from unmagic import fixture, use

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def round_tripped(tmp_path):
    src = tmp_path / "in.xlsx"
    out = tmp_path / "out.xlsx"
    write_simple_xlsx(src)
    wb = read_xlsx(src)
    write_xlsx(wb, out)
    yield read_xlsx(out)


@use(round_tripped)
def test_round_trip_preserves_sheets():
    wb = round_tripped()
    assert [s.name for s in wb.sheets] == ["Inputs", "Outputs"]


@use(round_tripped)
def test_round_trip_preserves_values():
    wb = round_tripped()
    assert wb.sheet("Inputs").get("B1").value == 0.04
    assert wb.sheet("Inputs").get("B2").value == 1_000_000


@use(round_tripped)
def test_round_trip_preserves_formula():
    wb = round_tripped()
    cell = wb.sheet("Outputs").get("B1")
    assert cell.formula == "=Inputs!B2 * (1 + Inputs!B1)"


@use(round_tripped)
def test_round_trip_preserves_named_range():
    wb = round_tripped()
    names = {nr.name: nr for nr in wb.named_ranges}
    assert names["growth_rate"].ref == "Inputs!$B$1"
```

- [ ] **Step 2: Run; confirm failures**

```bash
uv run pytest tests/test_xlsx_roundtrip.py -v
```

Expected: import errors (writer not yet implemented).

- [ ] **Step 3: Implement the writer**

`src/claudesheets/xlsx/writer.py`:

```python
"""Write a Workbook to an .xlsx file."""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from claudesheets.model.workbook import Workbook


def write_xlsx(wb: Workbook, path: Path) -> None:
    path = Path(path)
    out = openpyxl.Workbook()
    # openpyxl creates a default "Sheet"; remove it.
    default = out.active
    out.remove(default)

    for sheet in wb.sheets:
        ws = out.create_sheet(title=sheet.name)
        for addr, cell in sheet.cells.items():
            if cell.formula is not None:
                ws[addr] = cell.formula
            else:
                ws[addr] = cell.value

    for nr in wb.named_ranges:
        defn = DefinedName(name=nr.name, attr_text=nr.ref)
        if nr.scope == "workbook":
            out.defined_names[nr.name] = defn
        else:
            out[nr.sheet].defined_names[nr.name] = defn

    out.save(path)
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_xlsx_roundtrip.py -v
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/xlsx/writer.py tests/test_xlsx_roundtrip.py
git commit -m "xlsx: write values, formulas, sheets, and named ranges"
```

---

## Task 6: Cell formatting in the model

**Files:**
- Create: `src/claudesheets/model/format.py`
- Modify: `src/claudesheets/model/workbook.py` (add `formats` dict to `Sheet`)
- Modify: `src/claudesheets/model/cell.py` (already has `format_id`)
- Create: `tests/test_format_model.py`

- [ ] **Step 1: Write the failing test**

`tests/test_format_model.py`:

```python
from claudesheets.model.cell import Cell
from claudesheets.model.format import CellFormat, Font, Fill, Border, Side
from claudesheets.model.workbook import Sheet


def test_cellformat_default_is_empty():
    fmt = CellFormat()
    assert fmt.font is None
    assert fmt.fill is None
    assert fmt.border is None
    assert fmt.number_format is None


def test_font_fields():
    f = Font(name="Calibri", size=11.0, bold=True, italic=False, color="FF0000")
    assert f.bold and not f.italic


def test_fill_fields():
    f = Fill(color="FFFF00")
    assert f.color == "FFFF00"


def test_border_fields():
    b = Border(left=Side(style="thin", color="000000"))
    assert b.left.style == "thin"


def test_sheet_can_associate_format_with_cell():
    sh = Sheet(name="S")
    sh.formats["bold-red"] = CellFormat(
        font=Font(name="Calibri", size=11.0, bold=True, color="FF0000")
    )
    sh.set("A1", Cell(value="hello", format_id="bold-red"))
    assert sh.get("A1").format_id == "bold-red"
    assert sh.formats["bold-red"].font.bold is True
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement formats**

`src/claudesheets/model/format.py`:

```python
"""Cell-formatting model: font, fill, border, number format."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class Font:
    name: Optional[str] = None
    size: Optional[float] = None
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None  # "single", "double", or None
    color: Optional[str] = None  # hex RGB, e.g. "FF0000"


@dataclass(frozen=True)
class Fill:
    color: Optional[str] = None  # hex RGB; solid fill only in Plan 1


@dataclass(frozen=True)
class Side:
    style: Optional[str] = None  # "thin", "medium", "thick", "dashed", "dotted", "double"
    color: Optional[str] = None


@dataclass(frozen=True)
class Border:
    left: Optional[Side] = None
    right: Optional[Side] = None
    top: Optional[Side] = None
    bottom: Optional[Side] = None


@dataclass(frozen=True)
class CellFormat:
    font: Optional[Font] = None
    fill: Optional[Fill] = None
    border: Optional[Border] = None
    number_format: Optional[str] = None  # e.g. "0.00%", "#,##0.00"
```

Modify `src/claudesheets/model/workbook.py` — add `formats` to `Sheet`:

```python
@dataclass
class Sheet:
    name: str
    cells: Dict[str, Cell] = field(default_factory=dict)
    column_widths: Dict[str, float] = field(default_factory=dict)
    formats: Dict[str, "CellFormat"] = field(default_factory=dict)
    frozen_panes: Optional[str] = None
    # ... existing methods unchanged
```

(Add `from claudesheets.model.format import CellFormat` near the other imports.)

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_format_model.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/model tests/test_format_model.py
git commit -m "model: add CellFormat (font, fill, border, number format)"
```

---

## Task 7: xlsx reader/writer for formats and column widths

**Files:**
- Modify: `src/claudesheets/xlsx/reader.py`
- Modify: `src/claudesheets/xlsx/writer.py`
- Modify: `tests/fixtures/workbooks.py`
- Create: `tests/test_xlsx_format_roundtrip.py`

- [ ] **Step 1: Add a formatted-workbook fixture builder**

Append to `tests/fixtures/workbooks.py`:

```python
def write_formatted_xlsx(path: Path) -> None:
    """A workbook exercising fonts, fills, borders, number formats, column widths."""
    from openpyxl.styles import (
        Alignment,
        Border as XBorder,
        Fill as XFill,
        Font as XFont,
        PatternFill,
        Side as XSide,
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Format"

    ws["A1"] = "bold red"
    ws["A1"].font = XFont(name="Calibri", size=11, bold=True, color="FFFF0000")

    ws["B1"] = 0.1234
    ws["B1"].number_format = "0.00%"

    ws["C1"] = 1234.5
    ws["C1"].number_format = '#,##0.00'

    ws["D1"] = "filled"
    ws["D1"].fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")

    ws["E1"] = "boxed"
    side = XSide(style="thin", color="FF000000")
    ws["E1"].border = XBorder(left=side, right=side, top=side, bottom=side)

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 10

    wb.save(path)
```

- [ ] **Step 2: Write the failing test**

`tests/test_xlsx_format_roundtrip.py`:

```python
from unmagic import fixture, use

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_formatted_xlsx


@fixture
def round_tripped(tmp_path):
    src = tmp_path / "fmt.xlsx"
    out = tmp_path / "fmt-out.xlsx"
    write_formatted_xlsx(src)
    wb = read_xlsx(src)
    write_xlsx(wb, out)
    yield read_xlsx(out)


@use(round_tripped)
def test_font_round_trips():
    sheet = round_tripped().sheet("Format")
    cell = sheet.get("A1")
    fmt = sheet.formats[cell.format_id]
    assert fmt.font.bold is True
    assert fmt.font.color in ("FF0000", "FFFF0000")  # alpha-channel may or may not be preserved
    assert fmt.font.name == "Calibri"


@use(round_tripped)
def test_number_format_round_trips():
    sheet = round_tripped().sheet("Format")
    fmt_b = sheet.formats[sheet.get("B1").format_id]
    assert fmt_b.number_format == "0.00%"
    fmt_c = sheet.formats[sheet.get("C1").format_id]
    assert fmt_c.number_format == "#,##0.00"


@use(round_tripped)
def test_fill_round_trips():
    sheet = round_tripped().sheet("Format")
    fmt = sheet.formats[sheet.get("D1").format_id]
    assert fmt.fill.color in ("FFFF00", "FFFFFF00")


@use(round_tripped)
def test_border_round_trips():
    sheet = round_tripped().sheet("Format")
    fmt = sheet.formats[sheet.get("E1").format_id]
    assert fmt.border.left.style == "thin"
    assert fmt.border.right.style == "thin"
    assert fmt.border.top.style == "thin"
    assert fmt.border.bottom.style == "thin"


@use(round_tripped)
def test_column_widths_round_trip():
    sheet = round_tripped().sheet("Format")
    assert sheet.column_widths["A"] == 18
    assert sheet.column_widths["B"] == 10
```

- [ ] **Step 3: Run; confirm failures**

```bash
uv run pytest tests/test_xlsx_format_roundtrip.py -v
```

- [ ] **Step 4: Extend the reader**

Replace `src/claudesheets/xlsx/reader.py` with:

```python
"""Read an .xlsx file into a Workbook."""
from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.cell.cell import Cell as XCell

from claudesheets.model.cell import Cell
from claudesheets.model.format import Border, CellFormat, Fill, Font, Side
from claudesheets.model.workbook import NamedRange, Sheet, Workbook


def _normalise_color(c: object) -> Optional[str]:
    """openpyxl colours can be Color objects, theme refs, or hex strings.

    For Plan 1 we only carry hex RGB. Theme colours and indexed colours are
    dropped (None). Alpha-channel hex (8 chars) is normalised to 6-char RGB.
    """
    if c is None:
        return None
    rgb = getattr(c, "rgb", None) or (c if isinstance(c, str) else None)
    if not isinstance(rgb, str):
        return None
    if len(rgb) == 8:  # AARRGGBB
        return rgb[2:].upper()
    if len(rgb) == 6:
        return rgb.upper()
    return None


def _read_font(c: XCell) -> Optional[Font]:
    f = c.font
    if not f:
        return None
    has_anything = any([
        f.name, f.size, f.bold, f.italic, f.underline,
        _normalise_color(f.color),
    ])
    if not has_anything:
        return None
    return Font(
        name=f.name or None,
        size=float(f.size) if f.size is not None else None,
        bold=bool(f.bold),
        italic=bool(f.italic),
        underline=f.underline if f.underline in ("single", "double") else None,
        color=_normalise_color(f.color),
    )


def _read_fill(c: XCell) -> Optional[Fill]:
    fill = c.fill
    if not fill or fill.fill_type != "solid":
        return None
    color = _normalise_color(fill.fgColor)
    if not color:
        return None
    return Fill(color=color)


def _read_side(s) -> Optional[Side]:
    if not s or not s.style:
        return None
    return Side(style=s.style, color=_normalise_color(s.color))


def _read_border(c: XCell) -> Optional[Border]:
    b = c.border
    if not b:
        return None
    sides = {n: _read_side(getattr(b, n)) for n in ("left", "right", "top", "bottom")}
    if not any(sides.values()):
        return None
    return Border(**sides)


def _cell_format(c: XCell) -> Optional[CellFormat]:
    font = _read_font(c)
    fill = _read_fill(c)
    border = _read_border(c)
    nf = c.number_format if c.number_format and c.number_format != "General" else None
    if not any([font, fill, border, nf]):
        return None
    return CellFormat(font=font, fill=fill, border=border, number_format=nf)


def _format_id(fmt: CellFormat) -> str:
    """Stable, content-addressed id for a CellFormat."""
    h = hashlib.sha1(repr(fmt).encode("utf-8")).hexdigest()[:10]
    return f"f-{h}"


def read_xlsx(path: Path) -> Workbook:
    path = Path(path)
    src = openpyxl.load_workbook(path, data_only=False)
    wb = Workbook(name=path.stem)

    for ws in src.worksheets:
        sheet = Sheet(name=ws.title)
        for col_letter, dim in ws.column_dimensions.items():
            if dim.width is not None:
                sheet.column_widths[col_letter] = float(dim.width)
        for row in ws.iter_rows():
            for c in row:
                fmt = _cell_format(c)
                fmt_id = None
                if fmt is not None:
                    fmt_id = _format_id(fmt)
                    sheet.formats[fmt_id] = fmt

                if c.value is None and fmt_id is None:
                    continue
                if isinstance(c.value, str) and c.value.startswith("="):
                    sheet.set(c.coordinate, Cell(formula=c.value, format_id=fmt_id))
                elif c.value is None:
                    # blank but formatted
                    sheet.set(c.coordinate, Cell(format_id=fmt_id))
                else:
                    sheet.set(c.coordinate, Cell(value=c.value, format_id=fmt_id))
        wb.sheets.append(sheet)

    for name, defn in src.defined_names.items():
        wb.named_ranges.append(
            NamedRange(name=name, ref=defn.attr_text, scope="workbook")
        )
    for ws in src.worksheets:
        for name, defn in ws.defined_names.items():
            wb.named_ranges.append(
                NamedRange(name=name, ref=defn.attr_text, scope="sheet", sheet=ws.title)
            )

    return wb
```

- [ ] **Step 5: Extend the writer**

Replace `src/claudesheets/xlsx/writer.py` with:

```python
"""Write a Workbook to an .xlsx file."""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import (
    Border as XBorder,
    Color,
    Font as XFont,
    PatternFill,
    Side as XSide,
)
from openpyxl.workbook.defined_name import DefinedName

from claudesheets.model.format import CellFormat
from claudesheets.model.workbook import Workbook


def _xfont(font) -> Optional[XFont]:
    if font is None:
        return None
    color = Color(rgb=("FF" + font.color)) if font.color else None
    return XFont(
        name=font.name,
        size=font.size,
        bold=font.bold,
        italic=font.italic,
        underline=font.underline,
        color=color,
    )


def _xfill(fill) -> Optional[PatternFill]:
    if fill is None or not fill.color:
        return None
    return PatternFill(fill_type="solid", fgColor=Color(rgb=("FF" + fill.color)))


def _xside(side) -> Optional[XSide]:
    if side is None or not side.style:
        return None
    color = Color(rgb=("FF" + side.color)) if side.color else None
    return XSide(style=side.style, color=color)


def _xborder(border) -> Optional[XBorder]:
    if border is None:
        return None
    return XBorder(
        left=_xside(border.left),
        right=_xside(border.right),
        top=_xside(border.top),
        bottom=_xside(border.bottom),
    )


def _apply_format(cell, fmt: CellFormat) -> None:
    if fmt.font is not None:
        cell.font = _xfont(fmt.font)
    if fmt.fill is not None:
        cell.fill = _xfill(fmt.fill)
    if fmt.border is not None:
        cell.border = _xborder(fmt.border)
    if fmt.number_format is not None:
        cell.number_format = fmt.number_format


def write_xlsx(wb: Workbook, path: Path) -> None:
    path = Path(path)
    out = openpyxl.Workbook()
    out.remove(out.active)

    for sheet in wb.sheets:
        ws = out.create_sheet(title=sheet.name)

        for col, width in sheet.column_widths.items():
            ws.column_dimensions[col].width = width

        for addr, cell in sheet.cells.items():
            xc = ws[addr]
            if cell.formula is not None:
                xc.value = cell.formula
            elif cell.value is not None:
                xc.value = cell.value
            if cell.format_id and cell.format_id in sheet.formats:
                _apply_format(xc, sheet.formats[cell.format_id])

    for nr in wb.named_ranges:
        defn = DefinedName(name=nr.name, attr_text=nr.ref)
        if nr.scope == "workbook":
            out.defined_names[nr.name] = defn
        else:
            out[nr.sheet].defined_names[nr.name] = defn

    out.save(path)
```

- [ ] **Step 6: Run; confirm pass**

```bash
uv run pytest tests/test_xlsx_format_roundtrip.py tests/test_xlsx_roundtrip.py -v
```

Expected: all tests pass. If colour-comparison tests fail because of openpyxl alpha-channel quirks, the test already accepts both 6- and 8-char hex; if it still fails, narrow to the actual return value.

- [ ] **Step 7: Commit**

```bash
git add src/claudesheets/xlsx tests/fixtures/workbooks.py tests/test_xlsx_format_roundtrip.py
git commit -m "xlsx: round-trip fonts, fills, borders, number formats, column widths"
```

---

## Task 8: Data validation in the model and xlsx I/O

**Files:**
- Create: `src/claudesheets/model/validation.py`
- Modify: `src/claudesheets/model/workbook.py` (add `validations` to `Sheet`)
- Modify: `src/claudesheets/xlsx/reader.py`
- Modify: `src/claudesheets/xlsx/writer.py`
- Modify: `tests/fixtures/workbooks.py`
- Create: `tests/test_xlsx_validation_roundtrip.py`

- [ ] **Step 1: Add validation fixture**

Append to `tests/fixtures/workbooks.py`:

```python
def write_validation_xlsx(path: Path) -> None:
    from openpyxl.worksheet.datavalidation import DataValidation as XDV

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "V"

    dv_list = XDV(type="list", formula1='"yes,no,maybe"', allow_blank=True)
    dv_list.add("A1:A10")
    ws.add_data_validation(dv_list)

    dv_range = XDV(type="whole", operator="between", formula1=1, formula2=100)
    dv_range.add("B1:B10")
    ws.add_data_validation(dv_range)

    wb.save(path)
```

- [ ] **Step 2: Write the failing test**

`tests/test_xlsx_validation_roundtrip.py`:

```python
from unmagic import fixture, use

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_validation_xlsx


@fixture
def round_tripped(tmp_path):
    src = tmp_path / "v.xlsx"
    out = tmp_path / "v-out.xlsx"
    write_validation_xlsx(src)
    wb = read_xlsx(src)
    write_xlsx(wb, out)
    yield read_xlsx(out)


@use(round_tripped)
def test_list_validation_round_trips():
    sheet = round_tripped().sheet("V")
    matches = [v for v in sheet.validations if v.type == "list"]
    assert matches, "no list validation found"
    v = matches[0]
    assert "A1:A10" in v.ranges
    assert v.formula1 == '"yes,no,maybe"'


@use(round_tripped)
def test_whole_range_validation_round_trips():
    sheet = round_tripped().sheet("V")
    matches = [v for v in sheet.validations if v.type == "whole"]
    assert matches, "no whole-number validation found"
    v = matches[0]
    assert v.operator == "between"
    assert v.formula1 == 1 or v.formula1 == "1"
    assert v.formula2 == 100 or v.formula2 == "100"
    assert "B1:B10" in v.ranges
```

- [ ] **Step 3: Run; confirm failures**

- [ ] **Step 4: Implement validation model + I/O**

`src/claudesheets/model/validation.py`:

```python
"""Data validation rules."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class DataValidation:
    """A single data validation rule applied to one or more ranges.

    The `type` and `operator` strings match Excel's vocabulary
    ("list", "whole", "decimal", "date", "time", "textLength", "custom"
    for type; "between", "notBetween", "equal", "notEqual",
    "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"
    for operator).
    """
    type: str
    ranges: List[str] = field(default_factory=list)
    operator: Optional[str] = None
    formula1: Optional[str] = None
    formula2: Optional[str] = None
    allow_blank: bool = True
```

Add to `Sheet` in `src/claudesheets/model/workbook.py`:

```python
validations: List["DataValidation"] = field(default_factory=list)
```

(Add the appropriate import.)

Extend the reader's per-sheet block (insert after the `for col_letter` loop, before `for row in ws.iter_rows()`):

```python
for dv in ws.data_validations.dataValidation:
    sheet.validations.append(_read_validation(dv))
```

And add `_read_validation` to `src/claudesheets/xlsx/reader.py`:

```python
from claudesheets.model.validation import DataValidation


def _read_validation(dv) -> DataValidation:
    ranges = [str(r) for r in dv.sqref.ranges] if dv.sqref else []
    return DataValidation(
        type=dv.type,
        ranges=ranges,
        operator=dv.operator,
        formula1=dv.formula1,
        formula2=dv.formula2,
        allow_blank=bool(dv.allow_blank),
    )
```

Extend the writer's per-sheet block (after column widths, before cells):

```python
from openpyxl.worksheet.datavalidation import DataValidation as XDV

for v in sheet.validations:
    xdv = XDV(
        type=v.type,
        operator=v.operator,
        formula1=v.formula1,
        formula2=v.formula2,
        allow_blank=v.allow_blank,
    )
    for r in v.ranges:
        xdv.add(r)
    ws.add_data_validation(xdv)
```

(Move the `XDV` import to the top of the file.)

- [ ] **Step 5: Run; confirm pass**

```bash
uv run pytest tests/test_xlsx_validation_roundtrip.py -v
```

- [ ] **Step 6: Commit**

```bash
git add src/claudesheets/model src/claudesheets/xlsx tests/fixtures/workbooks.py tests/test_xlsx_validation_roundtrip.py
git commit -m "xlsx: round-trip data validation rules"
```

---

## Task 9: Markdown table parse/serialise

**Files:**
- Create: `src/claudesheets/source/__init__.py`
- Create: `src/claudesheets/source/markdown.py`
- Create: `tests/test_source_markdown.py`

The Markdown table format we're using: first row is the header `| (cell) | A | B | C |`; second row is the GitHub-style separator `| --- | --- | --- | --- |`; subsequent rows have the row number in the first column and cell content in subsequent columns. Empty cells are represented by an empty string between pipes. Formula cells contain the formula text starting with `=`. Numbers are written with `repr()` to preserve precision; strings are written verbatim.

This is a focused module, so we cover several behaviours in one task.

- [ ] **Step 1: Write the failing test**

`tests/test_source_markdown.py`:

```python
from claudesheets.model.cell import Cell
from claudesheets.model.workbook import Sheet
from claudesheets.source.markdown import dump_table, load_table


def test_round_trip_simple_values():
    sh = Sheet(name="S")
    sh.set("A1", Cell(value="hello"))
    sh.set("B1", Cell(value=42))
    sh.set("A2", Cell(value=3.14))
    sh.set("B2", Cell(value=True))
    text = dump_table(sh)
    sh2 = Sheet(name="S")
    load_table(sh2, text)
    assert sh2.get("A1").value == "hello"
    assert sh2.get("B1").value == 42
    assert sh2.get("A2").value == 3.14
    assert sh2.get("B2").value is True


def test_round_trip_formula():
    sh = Sheet(name="S")
    sh.set("A1", Cell(value=10))
    sh.set("B1", Cell(formula="=A1*2"))
    text = dump_table(sh)
    sh2 = Sheet(name="S")
    load_table(sh2, text)
    assert sh2.get("B1").formula == "=A1*2"


def test_round_trip_blanks_omitted():
    sh = Sheet(name="S")
    sh.set("A1", Cell(value=1))
    sh.set("C5", Cell(value=2))
    text = dump_table(sh)
    sh2 = Sheet(name="S")
    load_table(sh2, text)
    assert sh2.get("A1").value == 1
    assert sh2.get("B1").value is None
    assert sh2.get("C5").value == 2


def test_dump_includes_header_and_separator():
    sh = Sheet(name="S")
    sh.set("A1", Cell(value=1))
    text = dump_table(sh)
    lines = text.splitlines()
    assert lines[0].startswith("|")
    assert "A" in lines[0]
    assert set(lines[1].replace("|", "").strip()) <= {"-", " "}


def test_load_ignores_blank_rows_and_extra_whitespace():
    sh = Sheet(name="S")
    text = """\
| (cell) | A | B |
| --- | --- | --- |
|  1 |  hi |   |
|    |     |   |
|  3 |     | =A1+1 |
"""
    load_table(sh, text)
    assert sh.get("A1").value == "hi"
    assert sh.get("A3").value is None
    assert sh.get("B3").formula == "=A1+1"


def test_string_with_pipe_is_escaped():
    sh = Sheet(name="S")
    sh.set("A1", Cell(value="a | b"))
    text = dump_table(sh)
    sh2 = Sheet(name="S")
    load_table(sh2, text)
    assert sh2.get("A1").value == "a | b"
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement Markdown table I/O**

`src/claudesheets/source/__init__.py`: empty.

`src/claudesheets/source/markdown.py`:

```python
"""Parse and serialise a Sheet's value cells as a Markdown table.

Format:
- First row: `| (cell) | A | B | C | ...` where the second-onwards columns
  are the column letters. The first column header is the literal string
  `(cell)` and is reserved for the row index.
- Second row: GitHub-style separator `| --- | --- | --- | --- |`.
- Subsequent rows: `| <row-number> | <cell-A-content> | <cell-B-content> | ...`.

Cell content rules:
- Empty cell -> empty string.
- Strings are written verbatim, with `|` escaped as `\\|` and `\\` as `\\\\`.
- Numbers are written with repr() to preserve precision.
- Booleans are written as TRUE / FALSE (Excel's convention).
- Formula cells: the formula text, starting with `=`.

This module deliberately does NOT handle formats, validation, or
formula-keyed-by-name semantics — those live in the YAML sidecar.
"""
from __future__ import annotations

import re
from typing import List, Tuple

from openpyxl.utils import get_column_letter, column_index_from_string

from claudesheets.model.cell import Cell
from claudesheets.model.workbook import Sheet

_HEADER = "(cell)"


def _escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace("|", "\\|")


def _unescape(s: str) -> str:
    out: List[str] = []
    i = 0
    while i < len(s):
        c = s[i]
        if c == "\\" and i + 1 < len(s):
            out.append(s[i + 1])
            i += 2
            continue
        out.append(c)
        i += 1
    return "".join(out)


def _format_value(cell: Cell) -> str:
    if cell.formula is not None:
        return _escape(cell.formula)
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return repr(v)
    return _escape(str(v))


def _parse_value(raw: str) -> Cell:
    s = _unescape(raw.strip())
    if s == "":
        return Cell()
    if s.startswith("="):
        return Cell(formula=s)
    if s == "TRUE":
        return Cell(value=True)
    if s == "FALSE":
        return Cell(value=False)
    # Try int, then float, else string.
    try:
        if re.fullmatch(r"[+-]?\d+", s):
            return Cell(value=int(s))
    except ValueError:
        pass
    try:
        return Cell(value=float(s))
    except ValueError:
        return Cell(value=s)


def _parse_row(line: str) -> List[str]:
    # Split on unescaped `|`. Trim leading/trailing pipe.
    parts: List[str] = []
    buf: List[str] = []
    i = 0
    while i < len(line):
        c = line[i]
        if c == "\\" and i + 1 < len(line):
            buf.append(c)
            buf.append(line[i + 1])
            i += 2
            continue
        if c == "|":
            parts.append("".join(buf))
            buf = []
            i += 1
            continue
        buf.append(c)
        i += 1
    parts.append("".join(buf))
    # Trim the leading/trailing empty parts created by `|` at line edges.
    if parts and parts[0].strip() == "":
        parts = parts[1:]
    if parts and parts[-1].strip() == "":
        parts = parts[:-1]
    return parts


def dump_table(sheet: Sheet) -> str:
    if not sheet.cells:
        return f"| {_HEADER} |\n| --- |\n"

    max_row = 0
    max_col = 0
    for addr in sheet.cells:
        m = re.fullmatch(r"([A-Z]+)(\d+)", addr)
        if not m:
            raise ValueError(f"bad address: {addr!r}")
        col_letters, row_str = m.group(1), m.group(2)
        max_row = max(max_row, int(row_str))
        max_col = max(max_col, column_index_from_string(col_letters))

    headers = [_HEADER] + [get_column_letter(c) for c in range(1, max_col + 1)]
    sep = ["---"] * len(headers)

    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(sep) + " |",
    ]
    for r in range(1, max_row + 1):
        row_cells = [str(r)]
        for c in range(1, max_col + 1):
            addr = f"{get_column_letter(c)}{r}"
            row_cells.append(_format_value(sheet.get(addr)))
        lines.append("| " + " | ".join(row_cells) + " |")
    return "\n".join(lines) + "\n"


def load_table(sheet: Sheet, text: str) -> None:
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines) < 2:
        return  # header only or empty

    header = _parse_row(lines[0])
    if not header or header[0].strip() != _HEADER:
        raise ValueError("first column header must be (cell)")
    column_letters = [h.strip() for h in header[1:]]

    for line in lines[2:]:  # skip header + separator
        cells = _parse_row(line)
        if not cells:
            continue
        row_idx_raw = cells[0].strip()
        if not row_idx_raw or not row_idx_raw.isdigit():
            continue
        row_idx = int(row_idx_raw)
        for i, raw in enumerate(cells[1:]):
            if i >= len(column_letters):
                break
            cell = _parse_value(raw)
            if cell.is_blank:
                continue
            sheet.set(f"{column_letters[i]}{row_idx}", cell)
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_source_markdown.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/source tests/test_source_markdown.py
git commit -m "source: Markdown table parse/serialise"
```

---

## Task 10: YAML sidecar — formats, column widths, validation

**Files:**
- Create: `src/claudesheets/source/yaml_sidecar.py`
- Create: `tests/test_source_yaml.py`

Sidecar format (per sheet):

```yaml
column_widths:
  A: 18.0
  B: 10.0

formats:
  f-1234abcd:
    font: { name: Calibri, size: 11.0, bold: true, color: "FF0000" }
    fill: { color: "FFFF00" }
    border:
      left:   { style: thin, color: "000000" }
      right:  { style: thin, color: "000000" }
      top:    { style: thin, color: "000000" }
      bottom: { style: thin, color: "000000" }
    number_format: "0.00%"

cell_formats:
  A1: f-1234abcd
  B2: f-1234abcd

validations:
  - type: list
    ranges: ["A1:A10"]
    formula1: '"yes,no,maybe"'
    allow_blank: true
  - type: whole
    ranges: ["B1:B10"]
    operator: between
    formula1: 1
    formula2: 100
    allow_blank: true
```

Note: cell *values* and *formulas* live in the `.md`. The sidecar references cells by A1 address only.

- [ ] **Step 1: Write the failing test**

`tests/test_source_yaml.py`:

```python
from claudesheets.model.cell import Cell
from claudesheets.model.format import Border, CellFormat, Fill, Font, Side
from claudesheets.model.validation import DataValidation
from claudesheets.model.workbook import Sheet
from claudesheets.source.yaml_sidecar import dump_yaml, load_yaml


def _make_sheet() -> Sheet:
    sh = Sheet(name="S")
    sh.column_widths["A"] = 18.0
    sh.formats["fmt"] = CellFormat(
        font=Font(name="Calibri", size=11.0, bold=True, color="FF0000"),
        fill=Fill(color="FFFF00"),
        border=Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
        ),
        number_format="0.00%",
    )
    sh.set("A1", Cell(value=1, format_id="fmt"))
    sh.validations.append(DataValidation(
        type="list",
        ranges=["A1:A10"],
        formula1='"yes,no,maybe"',
        allow_blank=True,
    ))
    return sh


def test_yaml_round_trips_column_widths():
    sh = _make_sheet()
    text = dump_yaml(sh)
    sh2 = Sheet(name="S")
    sh2.set("A1", Cell(value=1))  # value loaded from md elsewhere
    load_yaml(sh2, text)
    assert sh2.column_widths["A"] == 18.0


def test_yaml_round_trips_formats_and_cell_formats():
    sh = _make_sheet()
    text = dump_yaml(sh)
    sh2 = Sheet(name="S")
    sh2.set("A1", Cell(value=1))
    load_yaml(sh2, text)
    assert "fmt" in sh2.formats
    fmt = sh2.formats["fmt"]
    assert fmt.font.bold is True
    assert fmt.font.color == "FF0000"
    assert fmt.fill.color == "FFFF00"
    assert fmt.number_format == "0.00%"
    assert sh2.get("A1").format_id == "fmt"


def test_yaml_round_trips_validations():
    sh = _make_sheet()
    text = dump_yaml(sh)
    sh2 = Sheet(name="S")
    load_yaml(sh2, text)
    assert len(sh2.validations) == 1
    v = sh2.validations[0]
    assert v.type == "list"
    assert v.ranges == ["A1:A10"]
    assert v.formula1 == '"yes,no,maybe"'
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement the sidecar**

`src/claudesheets/source/yaml_sidecar.py`:

```python
"""Read/write the per-sheet YAML sidecar.

The sidecar carries everything that doesn't fit cleanly in a Markdown
table: column widths, format definitions, the cell -> format mapping,
and data validation rules.

Cell values and formulas live in the .md, not here.
"""
from __future__ import annotations

import io
from dataclasses import asdict
from typing import Any, Dict, Optional

from ruamel.yaml import YAML

from claudesheets.model.cell import Cell
from claudesheets.model.format import Border, CellFormat, Fill, Font, Side
from claudesheets.model.validation import DataValidation
from claudesheets.model.workbook import Sheet

_yaml = YAML(typ="rt")
_yaml.indent(mapping=2, sequence=4, offset=2)
_yaml.preserve_quotes = True


def _font_to_dict(f: Font) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for k in ("name", "size", "bold", "italic", "underline", "color"):
        v = getattr(f, k)
        if v not in (None, False):
            out[k] = v
    return out


def _font_from_dict(d: Optional[Dict[str, Any]]) -> Optional[Font]:
    if not d:
        return None
    return Font(
        name=d.get("name"),
        size=d.get("size"),
        bold=bool(d.get("bold", False)),
        italic=bool(d.get("italic", False)),
        underline=d.get("underline"),
        color=d.get("color"),
    )


def _side_to_dict(s: Optional[Side]) -> Optional[Dict[str, Any]]:
    if s is None:
        return None
    out: Dict[str, Any] = {"style": s.style}
    if s.color:
        out["color"] = s.color
    return out


def _side_from_dict(d: Optional[Dict[str, Any]]) -> Optional[Side]:
    if not d:
        return None
    return Side(style=d.get("style"), color=d.get("color"))


def _fmt_to_dict(fmt: CellFormat) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    if fmt.font is not None:
        out["font"] = _font_to_dict(fmt.font)
    if fmt.fill is not None and fmt.fill.color:
        out["fill"] = {"color": fmt.fill.color}
    if fmt.border is not None:
        b: Dict[str, Any] = {}
        for n in ("left", "right", "top", "bottom"):
            sd = _side_to_dict(getattr(fmt.border, n))
            if sd is not None:
                b[n] = sd
        if b:
            out["border"] = b
    if fmt.number_format:
        out["number_format"] = fmt.number_format
    return out


def _fmt_from_dict(d: Dict[str, Any]) -> CellFormat:
    border_d = d.get("border")
    border = None
    if border_d:
        border = Border(
            left=_side_from_dict(border_d.get("left")),
            right=_side_from_dict(border_d.get("right")),
            top=_side_from_dict(border_d.get("top")),
            bottom=_side_from_dict(border_d.get("bottom")),
        )
    fill_d = d.get("fill")
    fill = Fill(color=fill_d.get("color")) if fill_d and fill_d.get("color") else None
    return CellFormat(
        font=_font_from_dict(d.get("font")),
        fill=fill,
        border=border,
        number_format=d.get("number_format"),
    )


def dump_yaml(sheet: Sheet) -> str:
    doc: Dict[str, Any] = {}

    if sheet.column_widths:
        doc["column_widths"] = dict(sheet.column_widths)

    if sheet.formats:
        doc["formats"] = {fid: _fmt_to_dict(f) for fid, f in sheet.formats.items()}

    cell_formats = {a: c.format_id for a, c in sheet.cells.items() if c.format_id}
    if cell_formats:
        doc["cell_formats"] = cell_formats

    if sheet.validations:
        doc["validations"] = [
            {k: v for k, v in asdict(dv).items() if v not in (None, [], False) or k == "ranges"}
            for dv in sheet.validations
        ]

    buf = io.StringIO()
    _yaml.dump(doc, buf)
    return buf.getvalue()


def load_yaml(sheet: Sheet, text: str) -> None:
    if not text.strip():
        return
    doc = _yaml.load(text) or {}

    for col, w in (doc.get("column_widths") or {}).items():
        sheet.column_widths[col] = float(w)

    for fid, d in (doc.get("formats") or {}).items():
        sheet.formats[fid] = _fmt_from_dict(dict(d))

    for addr, fid in (doc.get("cell_formats") or {}).items():
        existing = sheet.get(addr)
        sheet.set(addr, Cell(value=existing.value, formula=existing.formula, format_id=fid))

    for d in (doc.get("validations") or []):
        sheet.validations.append(DataValidation(
            type=d["type"],
            ranges=list(d.get("ranges") or []),
            operator=d.get("operator"),
            formula1=d.get("formula1"),
            formula2=d.get("formula2"),
            allow_blank=bool(d.get("allow_blank", True)),
        ))
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_source_yaml.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/source/yaml_sidecar.py tests/test_source_yaml.py
git commit -m "source: YAML sidecar for formats, widths, validation"
```

---

## Task 11: workbook.toml + claudesheets.toml read/write

**Files:**
- Create: `src/claudesheets/config.py`
- Create: `tests/test_config.py`

- [ ] **Step 1: Write the failing test**

`tests/test_config.py`:

```python
from claudesheets.config import (
    ProjectConfig,
    WorkbookManifest,
    dump_project,
    dump_workbook,
    load_project,
    load_workbook,
)
from claudesheets.model.workbook import NamedRange


def test_project_config_round_trips():
    cfg = ProjectConfig(name="my-model", calc_engine="libreoffice")
    text = dump_project(cfg)
    assert load_project(text) == cfg


def test_workbook_manifest_round_trips_basic():
    m = WorkbookManifest(name="my-model", sheets=["Inputs", "Outputs"], named_ranges=[])
    assert load_workbook(dump_workbook(m)) == m


def test_workbook_manifest_round_trips_named_ranges():
    m = WorkbookManifest(
        name="m",
        sheets=["S"],
        named_ranges=[
            NamedRange(name="x", scope="workbook", ref="S!$A$1"),
            NamedRange(name="y", scope="sheet", sheet="S", ref="$B$1"),
        ],
    )
    assert load_workbook(dump_workbook(m)) == m
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement config**

`src/claudesheets/config.py`:

```python
"""Read/write claudesheets.toml and workbook.toml."""
from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from typing import List

import tomli_w

from claudesheets.model.workbook import NamedRange


@dataclass(frozen=True)
class ProjectConfig:
    name: str
    calc_engine: str = "libreoffice"


@dataclass
class WorkbookManifest:
    name: str
    sheets: List[str] = field(default_factory=list)
    named_ranges: List[NamedRange] = field(default_factory=list)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, WorkbookManifest):
            return NotImplemented
        return (
            self.name == other.name
            and self.sheets == other.sheets
            and len(self.named_ranges) == len(other.named_ranges)
            and all(
                a.name == b.name and a.scope == b.scope
                and a.sheet == b.sheet and a.ref == b.ref
                for a, b in zip(self.named_ranges, other.named_ranges)
            )
        )


def dump_project(cfg: ProjectConfig) -> str:
    return tomli_w.dumps({
        "project": {"name": cfg.name},
        "build": {"calc_engine": cfg.calc_engine},
    })


def load_project(text: str) -> ProjectConfig:
    data = tomllib.loads(text)
    return ProjectConfig(
        name=data["project"]["name"],
        calc_engine=data.get("build", {}).get("calc_engine", "libreoffice"),
    )


def dump_workbook(m: WorkbookManifest) -> str:
    doc = {
        "workbook": {"name": m.name},
        "sheets": list(m.sheets),
        "named_ranges": [
            {"name": nr.name, "scope": nr.scope, **({"sheet": nr.sheet} if nr.sheet else {}),
             "ref": nr.ref}
            for nr in m.named_ranges
        ],
    }
    if not doc["named_ranges"]:
        del doc["named_ranges"]
    return tomli_w.dumps(doc)


def load_workbook(text: str) -> WorkbookManifest:
    data = tomllib.loads(text)
    nrs = [
        NamedRange(
            name=d["name"],
            ref=d["ref"],
            scope=d.get("scope", "workbook"),
            sheet=d.get("sheet"),
        )
        for d in data.get("named_ranges", [])
    ]
    return WorkbookManifest(
        name=data["workbook"]["name"],
        sheets=list(data.get("sheets", [])),
        named_ranges=nrs,
    )
```

Note that `tomli_w.dumps` produces TOML where the `sheets` key needs to live inside a section. Adjust if `dumps` rejects top-level arrays of plain strings; if it does, wrap as `{"workbook": {"name": ..., "sheets": [...]}}` and update `load_workbook` to read `data["workbook"]["sheets"]`.

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_config.py -v
```

If the `tomli_w` layout is rejected, fix per the note above and re-run.

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/config.py tests/test_config.py
git commit -m "config: read/write project and workbook manifests"
```

---

## Task 12: Source-directory reader/writer + Project class

**Files:**
- Create: `src/claudesheets/project.py`
- Create: `src/claudesheets/source/reader.py`
- Create: `src/claudesheets/source/writer.py`
- Create: `tests/test_source_roundtrip.py`

The naming convention for sheet files: each sheet `<name>` is stored as `sheets/<NN>_<slug>.md` and `sheets/<NN>_<slug>.yaml`, where `NN` is the position in workbook order (zero-padded to 2 digits) and `<slug>` is the lowercased sheet name with non-alphanumeric characters replaced by `_`. The original sheet name is recorded in `workbook.toml` (it's the entry in `sheets`); the filename is purely for ordering and human-friendliness.

- [ ] **Step 1: Write the failing test**

`tests/test_source_roundtrip.py`:

```python
from pathlib import Path

from unmagic import fixture, use

from claudesheets.model.cell import Cell
from claudesheets.model.format import CellFormat, Font
from claudesheets.model.validation import DataValidation
from claudesheets.model.workbook import NamedRange, Sheet, Workbook
from claudesheets.source.reader import read_source
from claudesheets.source.writer import write_source


def _make_workbook() -> Workbook:
    wb = Workbook(name="m")
    inputs = Sheet(name="Inputs")
    inputs.set("A1", Cell(value="growth"))
    inputs.set("B1", Cell(value=0.04))
    inputs.formats["bold"] = CellFormat(font=Font(name="Calibri", size=11.0, bold=True))
    inputs.set("A1", Cell(value="growth", format_id="bold"))
    inputs.column_widths["A"] = 18.0
    inputs.validations.append(DataValidation(
        type="list", ranges=["A1:A10"], formula1='"yes,no,maybe"', allow_blank=True,
    ))
    outputs = Sheet(name="Outputs")
    outputs.set("A1", Cell(value="rev"))
    outputs.set("B1", Cell(formula="=Inputs!B1*100"))
    wb.sheets = [inputs, outputs]
    wb.named_ranges.append(NamedRange(name="growth_rate", ref="Inputs!$B$1", scope="workbook"))
    return wb


@fixture
def project_dir(tmp_path):
    yield tmp_path / "proj"


@use(project_dir)
def test_round_trip_via_source_dir():
    wb = _make_workbook()
    write_source(wb, project_dir())
    wb2 = read_source(project_dir())
    assert [s.name for s in wb2.sheets] == ["Inputs", "Outputs"]
    assert wb2.sheet("Inputs").get("A1").value == "growth"
    assert wb2.sheet("Inputs").get("A1").format_id == "bold"
    assert wb2.sheet("Inputs").column_widths["A"] == 18.0
    assert wb2.sheet("Outputs").get("B1").formula == "=Inputs!B1*100"
    assert wb2.named_ranges[0].name == "growth_rate"
    assert wb2.named_ranges[0].ref == "Inputs!$B$1"
    v = wb2.sheet("Inputs").validations[0]
    assert v.ranges == ["A1:A10"]
    assert v.formula1 == '"yes,no,maybe"'
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement Project, source reader, source writer**

`src/claudesheets/project.py`:

```python
"""Locate and validate a claudesheets project directory."""
from __future__ import annotations

import re
from pathlib import Path

from claudesheets.exceptions import ProjectError


class Project:
    def __init__(self, root: Path):
        self.root = Path(root).resolve()

    @classmethod
    def open(cls, path: str | Path) -> "Project":
        root = Path(path).resolve()
        if not (root / "claudesheets.toml").is_file():
            raise ProjectError(f"Not a claudesheets project: {root}")
        return cls(root)

    @property
    def claudesheets_toml(self) -> Path:
        return self.root / "claudesheets.toml"

    @property
    def workbook_toml(self) -> Path:
        return self.root / "workbook.toml"

    @property
    def sheets_dir(self) -> Path:
        return self.root / "sheets"

    @property
    def data_dir(self) -> Path:
        return self.root / "data"

    @property
    def build_dir(self) -> Path:
        return self.root / "build"

    @property
    def cache_dir(self) -> Path:
        return self.root / ".claudesheets"

    @property
    def imports_dir(self) -> Path:
        return self.root / "imports"


def slugify(name: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_").lower()
    return s or "sheet"
```

`src/claudesheets/source/writer.py`:

```python
"""Write a Workbook to a claudesheets project directory."""
from __future__ import annotations

from pathlib import Path

from claudesheets.config import ProjectConfig, WorkbookManifest, dump_project, dump_workbook
from claudesheets.model.workbook import Workbook
from claudesheets.project import slugify
from claudesheets.source.markdown import dump_table
from claudesheets.source.yaml_sidecar import dump_yaml


def write_source(wb: Workbook, project_dir: Path) -> None:
    project_dir = Path(project_dir)
    project_dir.mkdir(parents=True, exist_ok=True)
    (project_dir / "sheets").mkdir(exist_ok=True)
    (project_dir / "data").mkdir(exist_ok=True)

    (project_dir / "claudesheets.toml").write_text(
        dump_project(ProjectConfig(name=wb.name))
    )
    (project_dir / "workbook.toml").write_text(
        dump_workbook(WorkbookManifest(
            name=wb.name,
            sheets=[s.name for s in wb.sheets],
            named_ranges=list(wb.named_ranges),
        ))
    )

    for i, sheet in enumerate(wb.sheets, start=1):
        stem = f"{i:02d}_{slugify(sheet.name)}"
        (project_dir / "sheets" / f"{stem}.md").write_text(dump_table(sheet))
        sidecar = dump_yaml(sheet)
        if sidecar.strip():
            (project_dir / "sheets" / f"{stem}.yaml").write_text(sidecar)
```

`src/claudesheets/source/reader.py`:

```python
"""Read a claudesheets project directory into a Workbook."""
from __future__ import annotations

from pathlib import Path

from claudesheets.config import load_project, load_workbook
from claudesheets.exceptions import ProjectError
from claudesheets.model.workbook import Sheet, Workbook
from claudesheets.project import slugify
from claudesheets.source.markdown import load_table
from claudesheets.source.yaml_sidecar import load_yaml


def read_source(project_dir: Path) -> Workbook:
    project_dir = Path(project_dir)
    cfg = load_project((project_dir / "claudesheets.toml").read_text())
    manifest = load_workbook((project_dir / "workbook.toml").read_text())

    wb = Workbook(name=manifest.name, named_ranges=list(manifest.named_ranges))
    for i, sheet_name in enumerate(manifest.sheets, start=1):
        stem = f"{i:02d}_{slugify(sheet_name)}"
        md_path = project_dir / "sheets" / f"{stem}.md"
        yaml_path = project_dir / "sheets" / f"{stem}.yaml"
        if not md_path.is_file():
            raise ProjectError(f"Missing sheet file: {md_path}")
        sheet = Sheet(name=sheet_name)
        load_table(sheet, md_path.read_text())
        if yaml_path.is_file():
            load_yaml(sheet, yaml_path.read_text())
        wb.sheets.append(sheet)
    return wb
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_source_roundtrip.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/project.py src/claudesheets/source tests/test_source_roundtrip.py
git commit -m "source: directory-level read/write of a Workbook"
```

---

## Task 13: Bulk CSV → cached SQLite

**Files:**
- Create: `src/claudesheets/bulk.py`
- Create: `tests/test_bulk.py`

For Plan 1, bulk CSVs are loaded into the cached SQLite as a build-time concern; sheets do not yet reference them via a `source:` declaration. This task delivers the load/build functionality so Plan 2 (when calc engine appears) can plug into it.

- [ ] **Step 1: Write the failing test**

`tests/test_bulk.py`:

```python
import sqlite3

from unmagic import fixture, use

from claudesheets.bulk import build_bulk_cache, table_name_for


@fixture
def project(tmp_path):
    proj = tmp_path / "p"
    (proj / "data").mkdir(parents=True)
    (proj / ".claudesheets").mkdir(parents=True)
    yield proj


@use(project)
def test_table_name_for_strips_extension():
    assert table_name_for("cpi_series.csv") == "cpi_series"
    assert table_name_for("Panel-Data.csv") == "panel_data"


@use(project)
def test_build_loads_csv_into_sqlite():
    p = project()
    (p / "data" / "series.csv").write_text("year,value\n2020,1.0\n2021,2.0\n2022,3.0\n")

    db_path = build_bulk_cache(p)
    assert db_path.exists()

    conn = sqlite3.connect(db_path)
    rows = conn.execute("SELECT year, value FROM series ORDER BY year").fetchall()
    assert rows == [("2020", "1.0"), ("2021", "2.0"), ("2022", "3.0")]


@use(project)
def test_build_skips_when_cache_is_newer():
    import time
    p = project()
    (p / "data" / "s.csv").write_text("a,b\n1,2\n")
    db_path = build_bulk_cache(p)
    first_mtime = db_path.stat().st_mtime

    time.sleep(0.01)
    build_bulk_cache(p)
    second_mtime = db_path.stat().st_mtime
    assert first_mtime == second_mtime, "cache rebuilt unnecessarily"


@use(project)
def test_build_rebuilds_when_csv_is_newer():
    import time
    p = project()
    (p / "data" / "s.csv").write_text("a,b\n1,2\n")
    db_path = build_bulk_cache(p)
    first_mtime = db_path.stat().st_mtime

    time.sleep(0.05)
    (p / "data" / "s.csv").write_text("a,b\n1,2\n3,4\n")
    build_bulk_cache(p)
    second_mtime = db_path.stat().st_mtime
    assert second_mtime > first_mtime
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement bulk cache**

`src/claudesheets/bulk.py`:

```python
"""Build the cached SQLite from data/*.csv."""
from __future__ import annotations

import csv
import re
import sqlite3
from pathlib import Path


def table_name_for(filename: str) -> str:
    stem = Path(filename).stem
    s = re.sub(r"[^A-Za-z0-9]+", "_", stem).strip("_").lower()
    return s or "_unnamed"


def _csvs_newer_than(csv_paths: list[Path], db_path: Path) -> bool:
    if not db_path.exists():
        return True
    db_mtime = db_path.stat().st_mtime
    return any(p.stat().st_mtime > db_mtime for p in csv_paths)


def build_bulk_cache(project_root: Path) -> Path:
    project_root = Path(project_root)
    data_dir = project_root / "data"
    cache_dir = project_root / ".claudesheets"
    cache_dir.mkdir(parents=True, exist_ok=True)
    db_path = cache_dir / "bulk.sqlite"

    csvs = sorted(data_dir.glob("*.csv")) if data_dir.is_dir() else []
    if not csvs:
        # No CSVs: ensure no stale db
        if db_path.exists():
            db_path.unlink()
        return db_path

    if not _csvs_newer_than(csvs, db_path):
        return db_path

    # Rebuild from scratch (deterministic).
    if db_path.exists():
        db_path.unlink()

    conn = sqlite3.connect(db_path)
    try:
        for csv_path in csvs:
            table = table_name_for(csv_path.name)
            with csv_path.open() as f:
                reader = csv.reader(f)
                header = next(reader, None)
                if not header:
                    continue
                # All columns typed TEXT in Plan 1; _schema.sql support is Plan 2/3.
                cols_sql = ", ".join(f'"{c}" TEXT' for c in header)
                conn.execute(f'CREATE TABLE "{table}" ({cols_sql})')
                placeholders = ", ".join(["?"] * len(header))
                conn.executemany(
                    f'INSERT INTO "{table}" VALUES ({placeholders})',
                    list(reader),
                )
        conn.commit()
    finally:
        conn.close()

    return db_path
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_bulk.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/bulk.py tests/test_bulk.py
git commit -m "bulk: load data/*.csv into a cached SQLite"
```

---

## Task 14: `claudesheets import` end-to-end

**Files:**
- Modify: `src/claudesheets/commands/import_cmd.py`
- Create: `tests/test_import.py`

- [ ] **Step 1: Write the failing test**

`tests/test_import.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from tests.fixtures.workbooks import write_simple_xlsx, write_formatted_xlsx


@fixture
def project(tmp_path):
    p = tmp_path / "proj"
    p.mkdir()
    (p / "claudesheets.toml").write_text('[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n')
    (p / "workbook.toml").write_text('[workbook]\nname = "x"\nsheets = []\n')
    (p / "sheets").mkdir()
    (p / "data").mkdir()
    yield p


@use(project)
def test_import_writes_source_files(tmp_path):
    src = tmp_path / "in.xlsx"
    write_simple_xlsx(src)
    runner = CliRunner()
    result = runner.invoke(main, ["import", str(src), "--project", str(project())])
    assert result.exit_code == 0, result.output
    assert (project() / "sheets" / "01_inputs.md").is_file()
    assert (project() / "sheets" / "02_outputs.md").is_file()
    text = (project() / "workbook.toml").read_text()
    assert "growth_rate" in text


@use(project)
def test_import_with_archive_copies_xlsx(tmp_path):
    src = tmp_path / "in.xlsx"
    write_simple_xlsx(src)
    runner = CliRunner()
    result = runner.invoke(
        main, ["import", str(src), "--archive", "--project", str(project())]
    )
    assert result.exit_code == 0, result.output
    archives = list((project() / "imports").glob("*.xlsx"))
    assert len(archives) == 1


@use(project)
def test_import_refuses_when_source_already_populated(tmp_path):
    src = tmp_path / "in.xlsx"
    write_simple_xlsx(src)
    # Pre-populate sheets/
    (project() / "sheets" / "01_existing.md").write_text("| (cell) | A |\n| --- | --- |\n")

    runner = CliRunner()
    result = runner.invoke(main, ["import", str(src), "--project", str(project())])
    assert result.exit_code != 0
    assert "non-empty" in result.output.lower() or "exist" in result.output.lower()


@use(project)
def test_import_round_trips_formatted_workbook(tmp_path):
    src = tmp_path / "fmt.xlsx"
    write_formatted_xlsx(src)
    runner = CliRunner()
    result = runner.invoke(main, ["import", str(src), "--project", str(project())])
    assert result.exit_code == 0, result.output
    yamls = list((project() / "sheets").glob("*.yaml"))
    assert yamls, "expected a YAML sidecar for the formatted sheet"
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement `import`**

`src/claudesheets/commands/import_cmd.py`:

```python
"""Implementation of `claudesheets import`."""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import click

from claudesheets.exceptions import ProjectError
from claudesheets.project import Project
from claudesheets.source.writer import write_source
from claudesheets.xlsx.reader import read_xlsx


def _detect_external_refs(xlsx_path: Path) -> list[str]:
    """Return the list of external workbook names referenced by formulas.

    openpyxl exposes external links via `wb._external_links`; for Plan 1 we
    take the conservative path of inspecting raw formulas after load.
    """
    import openpyxl
    src = openpyxl.load_workbook(xlsx_path, data_only=False)
    found: set[str] = set()
    for ws in src.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("=") and "[" in v and "]" in v:
                    found.add(v)
    return sorted(found)


def run(*, xlsx_path: str, project_path: str, archive: bool) -> None:
    xlsx = Path(xlsx_path).resolve()
    project_root = Path(project_path).resolve()

    try:
        Project.open(project_root)
    except ProjectError as e:
        raise click.ClickException(str(e))

    sheets_dir = project_root / "sheets"
    if any(sheets_dir.iterdir()):
        raise click.ClickException(
            f"sheets/ in {project_root} is non-empty; "
            "review-first re-import is deferred to a later release."
        )

    extrefs = _detect_external_refs(xlsx)
    if extrefs:
        raise click.ClickException(
            "Workbook contains external references; "
            "resolve them in Excel before importing.\n"
            "First few: " + ", ".join(extrefs[:3])
        )

    wb = read_xlsx(xlsx)
    write_source(wb, project_root)

    if archive:
        imports_dir = project_root / "imports"
        imports_dir.mkdir(exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%dT%H%M")
        shutil.copy2(xlsx, imports_dir / f"{ts}.xlsx")

    click.echo(f"Imported {xlsx} into {project_root}")
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_import.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/commands/import_cmd.py tests/test_import.py
git commit -m "import: ingest an .xlsx into source form"
```

---

## Task 15: `claudesheets build` end-to-end

**Files:**
- Modify: `src/claudesheets/commands/build_cmd.py`
- Create: `tests/test_build.py`

- [ ] **Step 1: Write the failing test**

`tests/test_build.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.xlsx.reader import read_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def imported_project(tmp_path):
    src = tmp_path / "in.xlsx"
    write_simple_xlsx(src)
    p = tmp_path / "proj"
    p.mkdir()
    (p / "claudesheets.toml").write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / "workbook.toml").write_text('[workbook]\nname = "in"\nsheets = []\n')
    (p / "sheets").mkdir()
    (p / "data").mkdir()

    runner = CliRunner()
    runner.invoke(main, ["import", str(src), "--project", str(p)])
    yield p


@use(imported_project)
def test_build_produces_xlsx():
    p = imported_project()
    runner = CliRunner()
    result = runner.invoke(main, ["build", "--project", str(p)])
    assert result.exit_code == 0, result.output
    out_xlsx = p / "build" / "in.xlsx"
    assert out_xlsx.is_file()


@use(imported_project)
def test_built_xlsx_has_same_sheets_and_values():
    p = imported_project()
    runner = CliRunner()
    runner.invoke(main, ["build", "--project", str(p)])
    wb = read_xlsx(p / "build" / "in.xlsx")
    assert [s.name for s in wb.sheets] == ["Inputs", "Outputs"]
    assert wb.sheet("Inputs").get("B1").value == 0.04
    assert wb.sheet("Outputs").get("B1").formula == "=Inputs!B2 * (1 + Inputs!B1)"
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement `build`**

`src/claudesheets/commands/build_cmd.py`:

```python
"""Implementation of `claudesheets build`."""
from __future__ import annotations

from pathlib import Path

import click

from claudesheets.bulk import build_bulk_cache
from claudesheets.config import load_project
from claudesheets.exceptions import ProjectError
from claudesheets.project import Project
from claudesheets.source.reader import read_source
from claudesheets.xlsx.writer import write_xlsx


def run(*, project_path: str, out_path: str | None) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    cfg = load_project(project.claudesheets_toml.read_text())

    wb = read_source(project.root)
    build_bulk_cache(project.root)

    project.build_dir.mkdir(parents=True, exist_ok=True)
    target = Path(out_path) if out_path else project.build_dir / f"{cfg.name}.xlsx"
    write_xlsx(wb, target)

    click.echo(f"Built {target}")
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_build.py -v
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/commands/build_cmd.py tests/test_build.py
git commit -m "build: compile source files into an .xlsx"
```

---

## Task 16: End-to-end fidelity test

**Files:**
- Create: `tests/test_end_to_end.py`

This task introduces no new code — it tightens the safety net by asserting that an `import → build` cycle produces an `.xlsx` semantically equivalent to the input across all Tier 1 features.

- [ ] **Step 1: Write the test**

`tests/test_end_to_end.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.model.workbook import Workbook
from claudesheets.xlsx.reader import read_xlsx
from tests.fixtures.workbooks import (
    write_formatted_xlsx,
    write_simple_xlsx,
    write_validation_xlsx,
)


def _round_trip(tmp_path: Path, builder) -> Workbook:
    src = tmp_path / "in.xlsx"
    builder(src)
    project = tmp_path / "proj"
    project.mkdir()
    (project / "claudesheets.toml").write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (project / "workbook.toml").write_text('[workbook]\nname = "in"\nsheets = []\n')
    (project / "sheets").mkdir()
    (project / "data").mkdir()

    runner = CliRunner()
    r = runner.invoke(main, ["import", str(src), "--project", str(project)])
    assert r.exit_code == 0, r.output
    r = runner.invoke(main, ["build", "--project", str(project)])
    assert r.exit_code == 0, r.output
    return read_xlsx(project / "build" / "in.xlsx")


def _assert_workbooks_equivalent(a: Workbook, b: Workbook) -> None:
    assert [s.name for s in a.sheets] == [s.name for s in b.sheets]
    for sa, sb in zip(a.sheets, b.sheets):
        assert sa.cells.keys() == sb.cells.keys(), f"cell sets differ on {sa.name}"
        for addr in sa.cells:
            ca, cb = sa.cells[addr], sb.cells[addr]
            assert ca.value == cb.value, f"value differs at {sa.name}!{addr}"
            assert ca.formula == cb.formula, f"formula differs at {sa.name}!{addr}"
        assert sa.column_widths == sb.column_widths, f"column widths differ on {sa.name}"
        # Formats compared by content (formats dict keys may differ in id).
        a_fmts = {ca.format_id and sa.formats.get(ca.format_id): addr
                  for addr, ca in sa.cells.items() if ca.format_id}
        b_fmts = {cb.format_id and sb.formats.get(cb.format_id): addr
                  for addr, cb in sb.cells.items() if cb.format_id}
        assert set(a_fmts.keys()) == set(b_fmts.keys()), f"format set differs on {sa.name}"
    assert {nr.name for nr in a.named_ranges} == {nr.name for nr in b.named_ranges}


@fixture
def tmp(tmp_path):
    yield tmp_path


@use(tmp)
def test_simple_workbook_round_trips_through_cli():
    out = _round_trip(tmp(), write_simple_xlsx)
    src = read_xlsx(_redo_simple(tmp()))
    _assert_workbooks_equivalent(src, out)


def _redo_simple(tmp_path: Path) -> Path:
    src = tmp_path / "ref.xlsx"
    write_simple_xlsx(src)
    return src


@use(tmp)
def test_formatted_workbook_round_trips_through_cli():
    out = _round_trip(tmp(), write_formatted_xlsx)
    ref = tmp() / "ref.xlsx"
    write_formatted_xlsx(ref)
    src = read_xlsx(ref)
    _assert_workbooks_equivalent(src, out)


@use(tmp)
def test_validation_workbook_round_trips_through_cli():
    out = _round_trip(tmp(), write_validation_xlsx)
    ref = tmp() / "ref.xlsx"
    write_validation_xlsx(ref)
    src = read_xlsx(ref)
    # Compare validations on each sheet.
    for sa, sb in zip(src.sheets, out.sheets):
        types_a = sorted(v.type for v in sa.validations)
        types_b = sorted(v.type for v in sb.validations)
        assert types_a == types_b
```

- [ ] **Step 2: Run; expect pass**

```bash
uv run pytest tests/test_end_to_end.py -v
```

If a test fails, the failure points to a specific feature whose round-trip is broken. Fix in the relevant module (xlsx reader/writer or source reader/writer), add a focused unit test if missing, then re-run.

- [ ] **Step 3: Run the full suite**

```bash
uv run pytest -v
```

Expected: all green.

- [ ] **Step 4: Commit**

```bash
git add tests/test_end_to_end.py
git commit -m "tests: end-to-end Tier 1 round-trip via CLI"
```

---

## Task 17: Documentation pass

**Files:**
- Modify: `README.md`
- Modify: `CLAUDE.md` (only if conventions changed during implementation)

- [ ] **Step 1: Update the README's Status section**

Change the Status section from "Pre-alpha. The design is settled; implementation has not started." to:

```
## Status

Plan 1 complete: Tier 1 round-trip foundation. `init`, `import`, `build`
work end-to-end for values, formulas, named ranges, basic formatting,
number formats, and data validation. Calc engine, tests, snapshot,
diff, and the escape-hatch import flow are coming in Plans 2–4.
```

- [ ] **Step 2: Verify everything still runs**

```bash
uv run pytest -v
uv run claudesheets --help
```

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: update status after Plan 1"
```

---

## Done criteria for Plan 1

All of these must be true to consider Plan 1 complete:

1. `uv run pytest -v` passes with all tests green.
2. `uv run claudesheets init <path>`, `claudesheets import <xlsx>`, and `claudesheets build` complete without error on a representative `.xlsx`.
3. The end-to-end test (`test_end_to_end.py`) passes for simple, formatted, and validation-bearing workbooks.
4. `import` correctly errors on workbooks with external references.
5. `import --archive` produces a timestamped copy in `imports/`.

## What this plan does **not** deliver (deferred)

- `recalc`, `test`, `snapshot`, `diff`, `check` commands — Plan 2/4
- Calc engine plugin and `claudesheets.testing` library — Plan 2
- Conditional formatting, comments, frozen panes, print areas, ListObject tables — Plan 3
- The escape-hatch re-import flow (hash detection, interactive merge) — Plan 4
- `--flatten` for external references — Plan 4
- MCP wrapper — Plan 5
