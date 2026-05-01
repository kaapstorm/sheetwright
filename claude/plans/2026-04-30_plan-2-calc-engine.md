# Plan 2 — Calc engine, testing library, recalc/test/snapshot

> **For agentic workers:** REQUIRED SUB-SKILL: Use
> `superpowers:subagent-driven-development` (recommended) or
> `superpowers:executing-plans` to implement this plan task-by-task.
> Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Wire up a swappable calc engine with a LibreOffice headless
implementation, expose a `claudesheets.testing.Model` API for pytest
tests, and ship the `recalc`, `test`, and `snapshot` commands.

**Architecture:**
- A `CalcEngine` ABC defines `evaluate(xlsx_path) -> dict[sheet, dict[addr, value]]`.
- `LibreOfficeEngine` implements it via `soffice --headless --calc
  --convert-to xlsx` to force a recalc, then reads cached values back
  with `openpyxl(..., data_only=True)`.
- A content-addressed calc cache (`.claudesheets/calc/<sha256>.json`)
  avoids re-running the engine when the built xlsx hasn't changed.
- `claudesheets.testing.Model` wraps a `Workbook` + `CalcEngine` and
  gives tests a `set/get/recalc` surface.
- `recalc` builds the xlsx, runs the engine, writes the cache.
- `test` shells out to `pytest` in `<project>/tests/`.
- `snapshot` reads the cache (running `recalc` first if missing) and
  diffs/updates `tests/snapshots/<workbook>.json`.

**Tech stack:** Python 3.11, openpyxl, click, subprocess, hashlib,
LibreOffice (`soffice` on `$PATH`), pytest, pytest-unmagic.

---

## File structure

New files:

```
src/claudesheets/calc/
├── __init__.py            # public API: get_calc_engine, CalcEngine
├── base.py                # CalcEngine ABC + CalcResult typing
├── libreoffice.py         # LibreOfficeEngine
└── cache.py               # hash + read/write JSON cache

src/claudesheets/testing/
├── __init__.py            # public API: Model, parse_address
└── model.py               # Model class

src/claudesheets/commands/
├── recalc_cmd.py          # implementation of `recalc`
├── test_cmd.py            # implementation of `test`
└── snapshot_cmd.py        # implementation of `snapshot`

src/claudesheets/snapshot.py  # snapshot dump/load/diff helpers

tests/
├── test_calc_libreoffice.py   # gated on `soffice` being available
├── test_calc_cache.py
├── test_testing_model.py
├── test_recalc_cmd.py
├── test_test_cmd.py
├── test_snapshot.py
├── test_snapshot_cmd.py
└── test_end_to_end_recalc.py
```

Modified files:

- `src/claudesheets/cli.py` — add `recalc`, `test`, `snapshot` subcommands
- `src/claudesheets/project.py` — add `tests_dir`, `calc_cache_dir`,
  `snapshots_dir` properties
- `src/claudesheets/commands/init_cmd.py` — scaffold `tests/` directory
  and a placeholder `tests/__init__.py`
- `pyproject.toml` — none expected (pytest already on dev deps; runtime
  pytest dependency added in Task 1 if we decide to ship it)
- `README.md` — Plan 2 status + brief calc-engine docs

## Conventions for every task

1. Single-quote strings (`quote-style = 'single'` in ruff config).
2. Type-hint every public function; `from __future__ import annotations`
   at the top of new modules.
3. Run `uv run ruff format <changed-files>` before each commit.
4. Run `uv run mypy src/` before any commit that touches types.
5. Run `uv run pytest -q` after each task; suite must stay green.
6. Each task ends with a small, focused commit.
7. Tests that require `soffice` on `$PATH` use a shared
   `requires_libreoffice` fixture (Task 2) that skips when missing.

---

## Task 0: Scaffold the new packages

**Files:**
- Create: `src/claudesheets/calc/__init__.py`
- Create: `src/claudesheets/testing/__init__.py`

- [ ] **Step 1: Create the empty packages**

`src/claudesheets/calc/__init__.py`:

```python
"""Calc engine plugin layer.

A calc engine evaluates a built `.xlsx` and returns the calculated
values for every cell. The default engine is LibreOffice headless.
"""
```

`src/claudesheets/testing/__init__.py`:

```python
"""Test-time helpers exposed to user pytest tests."""
```

- [ ] **Step 2: Verify**

```bash
uv run python -c "import claudesheets.calc, claudesheets.testing"
```

- [ ] **Step 3: Commit**

```bash
git add src/claudesheets/calc/__init__.py src/claudesheets/testing/__init__.py
git commit -m "calc: scaffold calc + testing packages"
```

---

## Task 1: `CalcEngine` interface

**Files:**
- Create: `src/claudesheets/calc/base.py`
- Modify: `src/claudesheets/calc/__init__.py`
- Create: `tests/test_calc_base.py`

- [ ] **Step 1: Write the failing test**

`tests/test_calc_base.py`:

```python
from pathlib import Path

import pytest

from claudesheets.calc import CalcEngine, CalcResult, get_calc_engine


def test_calc_engine_is_abstract():
    with pytest.raises(TypeError):
        CalcEngine()  # type: ignore[abstract]


def test_unknown_engine_raises():
    with pytest.raises(ValueError, match='unknown calc engine'):
        get_calc_engine('nonsense')


def test_calc_result_shape():
    result: CalcResult = {'Sheet1': {'A1': 1, 'B2': 'hello'}}
    assert result['Sheet1']['A1'] == 1


class _Recorder(CalcEngine):
    name = 'recorder'

    def evaluate(self, xlsx_path: Path) -> CalcResult:
        return {'Recorded': {'A1': str(xlsx_path)}}


def test_engine_subclass_evaluates(tmp_path):
    eng = _Recorder()
    out = eng.evaluate(tmp_path / 'x.xlsx')
    assert out == {'Recorded': {'A1': str(tmp_path / 'x.xlsx')}}
```

- [ ] **Step 2: Run; confirm failures**

```bash
uv run pytest tests/test_calc_base.py -v
```

- [ ] **Step 3: Implement**

`src/claudesheets/calc/base.py`:

```python
"""CalcEngine ABC and the registry."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict

from claudesheets.model.cell import CellValue

CalcResult = Dict[str, Dict[str, CellValue]]


class CalcEngine(ABC):
    """Abstract calc engine: evaluate a built .xlsx and return values."""

    name: str = ''

    @abstractmethod
    def evaluate(self, xlsx_path: Path) -> CalcResult:
        """Return calculated values keyed by sheet, then by A1 address."""
```

`src/claudesheets/calc/__init__.py`:

```python
"""Calc engine plugin layer.

A calc engine evaluates a built `.xlsx` and returns the calculated
values for every cell. The default engine is LibreOffice headless.
"""

from __future__ import annotations

from claudesheets.calc.base import CalcEngine, CalcResult


def get_calc_engine(name: str) -> CalcEngine:
    if name == 'libreoffice':
        from claudesheets.calc.libreoffice import LibreOfficeEngine

        return LibreOfficeEngine()
    raise ValueError(f'unknown calc engine: {name!r}')


__all__ = ['CalcEngine', 'CalcResult', 'get_calc_engine']
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_calc_base.py -v
uv run mypy src/
```

`get_calc_engine('libreoffice')` will fail until Task 2; that import
is lazy, so the tests above don't trigger it.

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/calc/ tests/test_calc_base.py
git commit -m "calc: add CalcEngine ABC and registry"
```

---

## Task 2: `LibreOfficeEngine`

**Files:**
- Create: `src/claudesheets/calc/libreoffice.py`
- Create: `src/claudesheets/testing/fixtures.py`
- Create: `tests/test_calc_libreoffice.py`

The engine recalculates by re-saving the workbook with LibreOffice,
which forces formula evaluation, then reads `data_only=True` to
extract calculated values.

- [ ] **Step 1: Write the shared `requires_libreoffice` fixture**

`src/claudesheets/testing/fixtures.py`:

```python
"""Shared pytest-unmagic fixtures for claudesheets tests.

Living under `claudesheets.testing` so user projects can reuse them.
"""

from __future__ import annotations

import shutil

import pytest
from unmagic import fixture


@fixture
def requires_libreoffice():
    """Skip the test if `soffice` is not on $PATH."""
    if shutil.which('soffice') is None:
        pytest.skip('LibreOffice (soffice) not on $PATH')
    yield
```

- [ ] **Step 2: Write the failing test**

`tests/test_calc_libreoffice.py`:

```python
from pathlib import Path

from unmagic import use

from claudesheets.calc.libreoffice import LibreOfficeEngine
from claudesheets.testing.fixtures import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@use(requires_libreoffice)
def test_libreoffice_engine_returns_calculated_values(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    eng = LibreOfficeEngine()
    out = eng.evaluate(src)
    # Inputs!B1 = 0.04 (literal), Inputs!B2 = 1_000_000 (literal),
    # Outputs!B1 = =Inputs!B2 * (1 + Inputs!B1) = 1_040_000.
    assert out['Inputs']['B1'] == 0.04
    assert out['Inputs']['B2'] == 1_000_000
    assert out['Outputs']['B1'] == 1_040_000


@use(requires_libreoffice)
def test_libreoffice_engine_propagates_string_values(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    out = LibreOfficeEngine().evaluate(src)
    assert out['Inputs']['A1'] == 'growth_rate'
```

- [ ] **Step 3: Run; confirm failures**

```bash
uv run pytest tests/test_calc_libreoffice.py -v
```

- [ ] **Step 4: Implement**

`src/claudesheets/calc/libreoffice.py`:

```python
"""LibreOffice headless calc engine.

Strategy: copy the input xlsx into a fresh temp dir, run
`soffice --headless --calc --convert-to xlsx --outdir <tmp> <input>`
which forces LibreOffice to recalculate every formula and persist
cached values, then load the converted file with openpyxl in
`data_only=True` mode and read every cell.

We use `-env:UserInstallation=file://<tmp>/profile` to give each run
its own profile directory so concurrent invocations do not block each
other on LibreOffice's user-profile lock.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl

from claudesheets.calc.base import CalcEngine, CalcResult


class LibreOfficeError(RuntimeError):
    """soffice exited non-zero or produced no output."""


class LibreOfficeEngine(CalcEngine):
    name = 'libreoffice'

    def __init__(self, soffice: str = 'soffice', timeout: float = 120.0):
        self.soffice = soffice
        self.timeout = timeout

    def evaluate(self, xlsx_path: Path) -> CalcResult:
        xlsx_path = Path(xlsx_path).resolve()
        with tempfile.TemporaryDirectory(prefix='cshs-calc-') as td_str:
            td = Path(td_str)
            profile = td / 'profile'
            outdir = td / 'out'
            outdir.mkdir()

            # soffice --convert-to writes to <outdir>/<stem>.xlsx
            cmd = [
                self.soffice,
                '--headless',
                '--calc',
                f'-env:UserInstallation=file://{profile}',
                '--convert-to',
                'xlsx',
                '--outdir',
                str(outdir),
                str(xlsx_path),
            ]
            try:
                proc = subprocess.run(
                    cmd,
                    capture_output=True,
                    timeout=self.timeout,
                    check=False,
                )
            except FileNotFoundError as e:
                raise LibreOfficeError(
                    f'{self.soffice} not on $PATH'
                ) from e

            if proc.returncode != 0:
                raise LibreOfficeError(
                    f'soffice exited {proc.returncode}: '
                    f'{proc.stderr.decode("utf-8", "replace").strip()}'
                )

            converted = outdir / xlsx_path.name
            if not converted.is_file():
                # soffice sometimes appends .xlsx to a stem; fall back.
                candidates = list(outdir.glob('*.xlsx'))
                if len(candidates) != 1:
                    raise LibreOfficeError(
                        f'expected 1 xlsx in {outdir}, found {candidates}'
                    )
                converted = candidates[0]

            return _read_calculated(converted)


def _read_calculated(xlsx_path: Path) -> CalcResult:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    out: CalcResult = {}
    for ws in wb.worksheets:
        sheet: dict[str, object] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                sheet[cell.coordinate] = cell.value
        out[ws.title] = sheet  # type: ignore[assignment]
    return out
```

Note: `_read_calculated` uses a local mutable `dict[str, object]` then
assigns to the typed `CalcResult`; mypy is satisfied via the cast in
the assignment ignore. If mypy disagrees, change the local type to
`Dict[str, CellValue]` and import `CellValue`.

- [ ] **Step 5: Run; confirm pass**

```bash
uv run pytest tests/test_calc_libreoffice.py -v
uv run mypy src/
```

- [ ] **Step 6: Commit**

```bash
git add src/claudesheets/calc/libreoffice.py \
        src/claudesheets/testing/fixtures.py \
        tests/test_calc_libreoffice.py
git commit -m "calc: LibreOffice headless engine"
```

---

## Task 3: Calc cache

**Files:**
- Create: `src/claudesheets/calc/cache.py`
- Create: `tests/test_calc_cache.py`
- Modify: `src/claudesheets/project.py`

- [ ] **Step 1: Add the project property**

In `src/claudesheets/project.py`, after `cache_dir`, add:

```python
    @property
    def calc_cache_dir(self) -> Path:
        return self.cache_dir / 'calc'
```

- [ ] **Step 2: Write the failing test**

`tests/test_calc_cache.py`:

```python
from pathlib import Path

from claudesheets.calc.cache import (
    cache_path_for,
    hash_xlsx,
    read_cached,
    write_cached,
)


def test_hash_xlsx_is_stable(tmp_path: Path):
    p = tmp_path / 'a.xlsx'
    p.write_bytes(b'hello')
    h1 = hash_xlsx(p)
    h2 = hash_xlsx(p)
    assert h1 == h2
    assert len(h1) == 64  # sha256 hex


def test_hash_xlsx_changes_with_content(tmp_path: Path):
    a = tmp_path / 'a.xlsx'
    b = tmp_path / 'b.xlsx'
    a.write_bytes(b'hello')
    b.write_bytes(b'world')
    assert hash_xlsx(a) != hash_xlsx(b)


def test_cache_round_trips(tmp_path: Path):
    cache_dir = tmp_path / 'cache'
    result = {'Sheet1': {'A1': 42, 'B2': 'x'}}
    write_cached(cache_dir, 'abc', result)
    assert read_cached(cache_dir, 'abc') == result


def test_cache_miss_returns_none(tmp_path: Path):
    assert read_cached(tmp_path / 'cache', 'never') is None


def test_cache_path_layout(tmp_path: Path):
    p = cache_path_for(tmp_path / 'cache', 'deadbeef')
    assert p == tmp_path / 'cache' / 'deadbeef.json'
```

- [ ] **Step 3: Run; confirm failures**

- [ ] **Step 4: Implement**

`src/claudesheets/calc/cache.py`:

```python
"""Content-addressed calc cache.

The cache key is the SHA-256 of the built xlsx bytes. Hits avoid
running the calc engine entirely.
"""

from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Optional

from claudesheets.calc.base import CalcResult


def hash_xlsx(xlsx_path: Path) -> str:
    h = hashlib.sha256()
    with open(xlsx_path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def cache_path_for(cache_dir: Path, key: str) -> Path:
    return cache_dir / f'{key}.json'


def read_cached(cache_dir: Path, key: str) -> Optional[CalcResult]:
    p = cache_path_for(cache_dir, key)
    if not p.is_file():
        return None
    data = json.loads(p.read_text())
    return data['result']


def write_cached(
    cache_dir: Path, key: str, result: CalcResult
) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    p = cache_path_for(cache_dir, key)
    p.write_text(
        json.dumps(
            {
                'key': key,
                'computed_at': datetime.utcnow().isoformat() + 'Z',
                'result': result,
            },
            indent=2,
            sort_keys=True,
            default=_json_default,
        )
    )
    return p


def _json_default(obj: object) -> object:
    if isinstance(obj, datetime):
        return obj.isoformat() + 'Z'
    raise TypeError(f'not JSON serializable: {type(obj).__name__}')
```

- [ ] **Step 5: Run; confirm pass**

- [ ] **Step 6: Commit**

```bash
git add src/claudesheets/calc/cache.py src/claudesheets/project.py \
        tests/test_calc_cache.py
git commit -m "calc: content-addressed calc cache"
```

---

## Task 4: `claudesheets recalc` command

**Files:**
- Create: `src/claudesheets/commands/recalc_cmd.py`
- Modify: `src/claudesheets/cli.py`
- Create: `tests/test_recalc_cmd.py`

- [ ] **Step 1: Wire CLI**

In `src/claudesheets/cli.py`, after `build_cmd`:

```python
@main.command('recalc')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
)
@click.option(
    '--force',
    is_flag=True,
    help='Ignore the cache and re-run the calc engine.',
)
def recalc_cmd(project_path: str, force: bool) -> None:
    """Run the calc engine and cache calculated values."""
    from claudesheets.commands.recalc_cmd import run

    run(project_path=project_path, force=force)
```

- [ ] **Step 2: Write the failing test**

`tests/test_recalc_cmd.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.testing.fixtures import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def imported(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text(
        '[workbook]\nname = "in"\nsheets = []\n'
    )
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    runner = CliRunner()
    r = runner.invoke(main, ['import', str(src), '--project', str(p)])
    assert r.exit_code == 0, r.output
    r = runner.invoke(main, ['build', '--project', str(p)])
    assert r.exit_code == 0, r.output
    yield p


@use(imported, requires_libreoffice)
def test_recalc_writes_cache_entry():
    p = imported()
    runner = CliRunner()
    r = runner.invoke(main, ['recalc', '--project', str(p)])
    assert r.exit_code == 0, r.output
    cache_files = list((p / '.claudesheets' / 'calc').glob('*.json'))
    assert len(cache_files) == 1


@use(imported, requires_libreoffice)
def test_recalc_is_idempotent_uses_cache():
    p = imported()
    runner = CliRunner()
    runner.invoke(main, ['recalc', '--project', str(p)])
    out = runner.invoke(main, ['recalc', '--project', str(p)])
    assert out.exit_code == 0
    assert 'cache hit' in out.output.lower()


@use(imported, requires_libreoffice)
def test_recalc_force_rebuilds_cache():
    p = imported()
    runner = CliRunner()
    runner.invoke(main, ['recalc', '--project', str(p)])
    out = runner.invoke(
        main, ['recalc', '--project', str(p), '--force']
    )
    assert out.exit_code == 0
    assert 'cache hit' not in out.output.lower()
```

- [ ] **Step 3: Run; confirm failures**

- [ ] **Step 4: Implement**

`src/claudesheets/commands/recalc_cmd.py`:

```python
"""Implementation of `claudesheets recalc`."""

from __future__ import annotations

from pathlib import Path

import click

from claudesheets.calc import get_calc_engine
from claudesheets.calc.cache import (
    hash_xlsx,
    read_cached,
    write_cached,
)
from claudesheets.config import load_project
from claudesheets.exceptions import ProjectError
from claudesheets.project import Project


def run(*, project_path: str, force: bool) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    cfg = load_project(project.claudesheets_toml.read_text())
    built = project.build_dir / f'{cfg.name}.xlsx'
    if not built.is_file():
        raise click.ClickException(
            f'No built xlsx at {built}. Run `claudesheets build` first.'
        )

    key = hash_xlsx(built)
    if not force:
        cached = read_cached(project.calc_cache_dir, key)
        if cached is not None:
            click.echo(f'cache hit: {key[:12]}')
            return

    engine = get_calc_engine(cfg.calc_engine)
    result = engine.evaluate(built)
    path = write_cached(project.calc_cache_dir, key, result)
    click.echo(f'recalculated: {path}')
```

- [ ] **Step 5: Run; confirm pass**

```bash
uv run pytest tests/test_recalc_cmd.py -v
uv run mypy src/
```

- [ ] **Step 6: Commit**

```bash
git add src/claudesheets/commands/recalc_cmd.py src/claudesheets/cli.py \
        tests/test_recalc_cmd.py
git commit -m "recalc: run calc engine and cache values"
```

---

## Task 5: Address parsing helpers

**Files:**
- Create: `src/claudesheets/testing/addresses.py`
- Create: `tests/test_addresses.py`

- [ ] **Step 1: Write the failing test**

`tests/test_addresses.py`:

```python
import pytest

from claudesheets.model.workbook import NamedRange, Sheet, Workbook
from claudesheets.testing.addresses import parse_address


def _wb_with_named(name: str, ref: str) -> Workbook:
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    wb.named_ranges.append(NamedRange(name=name, ref=ref))
    return wb


def test_parse_qualified_a1():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'Inputs!B1') == ('Inputs', 'B1')


def test_parse_named_range_resolves_to_address():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'growth_rate') == ('Inputs', 'B1')


def test_parse_unknown_name_raises():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    with pytest.raises(KeyError, match='unknown'):
        parse_address(wb, 'no_such_name')


def test_parse_bare_a1_without_sheet_raises():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    with pytest.raises(ValueError, match='must include sheet'):
        parse_address(wb, 'B1')


def test_parse_strips_dollar_signs():
    wb = _wb_with_named('growth_rate', 'Inputs!$B$1')
    assert parse_address(wb, 'Inputs!$B$1') == ('Inputs', 'B1')


def test_parse_quoted_sheet_name():
    wb = Workbook(name='x', sheets=[Sheet(name='My Sheet')])
    assert parse_address(wb, "'My Sheet'!A1") == ('My Sheet', 'A1')
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement**

`src/claudesheets/testing/addresses.py`:

```python
"""Resolve user-facing addresses to (sheet, A1) pairs.

Accepts:
- `Sheet!A1` or `'Sheet With Space'!A1` (Excel-style sheet-qualified)
- `name` (a workbook-scoped named range whose `ref` is a single cell)

Bare A1 (no sheet) is rejected: in this codebase tests always say
*which* sheet they mean.
"""

from __future__ import annotations

import re
from typing import Tuple

from claudesheets.model.workbook import Workbook


_QUOTED = re.compile(r"^'((?:[^']|'')+)'!(.+)$")
_UNQUOTED = re.compile(r'^([^!]+)!(.+)$')


def parse_address(wb: Workbook, address: str) -> Tuple[str, str]:
    if '!' in address:
        m = _QUOTED.match(address) or _UNQUOTED.match(address)
        if not m:
            raise ValueError(f'unparseable address: {address!r}')
        sheet, addr = m.group(1), m.group(2)
        sheet = sheet.replace("''", "'")
        return sheet, _strip_dollars(addr)

    for nr in wb.named_ranges:
        if nr.name == address:
            return parse_address(wb, nr.ref)

    if re.fullmatch(r'\$?[A-Za-z]+\$?\d+', address):
        raise ValueError(
            f'address {address!r} must include sheet (e.g. Sheet!A1)'
        )
    raise KeyError(f'unknown name: {address!r}')


def _strip_dollars(a1: str) -> str:
    return a1.replace('$', '')
```

- [ ] **Step 4: Run; confirm pass**

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/testing/addresses.py tests/test_addresses.py
git commit -m "testing: parse Sheet!A1 and named-range addresses"
```

---

## Task 6: `claudesheets.testing.Model`

**Files:**
- Create: `src/claudesheets/testing/model.py`
- Modify: `src/claudesheets/testing/__init__.py`
- Create: `tests/test_testing_model.py`

- [ ] **Step 1: Write the failing test**

`tests/test_testing_model.py`:

```python
from pathlib import Path

from unmagic import fixture, use

from claudesheets.testing import Model
from claudesheets.testing.fixtures import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def project(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text(
        '[workbook]\nname = "in"\nsheets = []\n'
    )
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    from click.testing import CliRunner
    from claudesheets.cli import main

    r = CliRunner().invoke(
        main, ['import', str(src), '--project', str(p)]
    )
    assert r.exit_code == 0
    yield p


@use(project, requires_libreoffice)
def test_model_get_returns_calculated_value():
    m = Model.open(project())
    assert m.get('Outputs!B1') == 1_040_000


@use(project, requires_libreoffice)
def test_model_set_invalidates_cache_and_recalculates():
    m = Model.open(project())
    m.set('Inputs!B1', 0.10)
    assert m.get('Outputs!B1') == 1_100_000


@use(project, requires_libreoffice)
def test_model_set_resolves_named_range():
    m = Model.open(project())
    m.set('growth_rate', 0.20)
    assert m.get('Outputs!B1') == 1_200_000


@use(project, requires_libreoffice)
def test_model_get_literal_value():
    m = Model.open(project())
    assert m.get('Inputs!B2') == 1_000_000
```

- [ ] **Step 2: Run; confirm failures**

- [ ] **Step 3: Implement**

`src/claudesheets/testing/model.py`:

```python
"""User-facing test model: set/get/recalc over a calc engine."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

from claudesheets.calc import CalcEngine, CalcResult, get_calc_engine
from claudesheets.config import load_project
from claudesheets.model.cell import Cell, CellValue
from claudesheets.model.workbook import Workbook
from claudesheets.project import Project
from claudesheets.source.reader import read_source
from claudesheets.testing.addresses import parse_address
from claudesheets.xlsx.writer import write_xlsx


class Model:
    """A workbook held in memory plus the calc engine that evaluates it.

    `set` mutates the in-memory workbook and invalidates calculated
    values. `get` triggers a recalc on demand and returns the
    calculated value (or the literal value for non-formula cells).
    """

    def __init__(self, wb: Workbook, engine: CalcEngine):
        self._wb = wb
        self._engine = engine
        self._calculated: Optional[CalcResult] = None

    @classmethod
    def open(cls, project_path: str | Path) -> 'Model':
        project = Project.open(project_path)
        cfg = load_project(project.claudesheets_toml.read_text())
        wb = read_source(project.root)
        return cls(wb, get_calc_engine(cfg.calc_engine))

    @property
    def workbook(self) -> Workbook:
        return self._wb

    def set(self, address: str, value: CellValue) -> None:
        sheet_name, addr = parse_address(self._wb, address)
        sheet = self._wb.sheet(sheet_name)
        existing = sheet.get(addr)
        sheet.set(addr, Cell(value=value, format_id=existing.format_id))
        self._calculated = None

    def get(self, address: str) -> CellValue:
        sheet_name, addr = parse_address(self._wb, address)
        if self._calculated is None:
            self.recalc()
        assert self._calculated is not None  # for mypy
        sheet = self._calculated.get(sheet_name, {})
        if addr in sheet:
            return sheet[addr]
        # Fall back to the literal value for cells the engine omitted.
        return self._wb.sheet(sheet_name).get(addr).value

    def recalc(self) -> None:
        with tempfile.TemporaryDirectory(prefix='cshs-model-') as td:
            xlsx = Path(td) / 'model.xlsx'
            write_xlsx(self._wb, xlsx)
            self._calculated = self._engine.evaluate(xlsx)
```

`src/claudesheets/testing/__init__.py`:

```python
"""Test-time helpers exposed to user pytest tests."""

from __future__ import annotations

from claudesheets.testing.addresses import parse_address
from claudesheets.testing.model import Model

__all__ = ['Model', 'parse_address']
```

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_testing_model.py -v
uv run mypy src/
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/testing/ tests/test_testing_model.py
git commit -m "testing: Model.set/get/recalc over a calc engine"
```

---

## Task 7: `claudesheets test` command

**Files:**
- Create: `src/claudesheets/commands/test_cmd.py`
- Modify: `src/claudesheets/cli.py`
- Modify: `src/claudesheets/project.py` (add `tests_dir`)
- Modify: `src/claudesheets/commands/init_cmd.py` (scaffold `tests/`)
- Create: `tests/test_test_cmd.py`

- [ ] **Step 1: Add `tests_dir` to `Project`**

In `src/claudesheets/project.py`:

```python
    @property
    def tests_dir(self) -> Path:
        return self.root / 'tests'
```

- [ ] **Step 2: Make `init` create `tests/`**

In `src/claudesheets/commands/init_cmd.py`, after the existing
`(project / 'data').mkdir()` line, add:

```python
    (project / 'tests').mkdir()
    (project / 'tests' / '__init__.py').write_text('')
```

Then extend `tests/test_init.py::test_init_creates_project_skeleton`
to also assert:

```python
    assert (project / 'tests').is_dir()
    assert (project / 'tests' / '__init__.py').is_file()
```

- [ ] **Step 3: Wire CLI**

In `src/claudesheets/cli.py`, add:

```python
@main.command(
    'test',
    context_settings={'ignore_unknown_options': True},
)
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
)
@click.argument('pytest_args', nargs=-1, type=click.UNPROCESSED)
def test_cmd(project_path: str, pytest_args: tuple[str, ...]) -> None:
    """Run the project's pytest tests."""
    from claudesheets.commands.test_cmd import run

    run(project_path=project_path, pytest_args=list(pytest_args))
```

- [ ] **Step 4: Write the failing test**

`tests/test_test_cmd.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main


@fixture
def project_with_tests(tmp_path: Path):
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text(
        '[workbook]\nname = "x"\nsheets = []\n'
    )
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    (p / 'tests').mkdir()
    (p / 'tests' / '__init__.py').write_text('')
    (p / 'tests' / 'test_simple.py').write_text(
        'def test_passes():\n    assert 1 + 1 == 2\n'
    )
    yield p


@use(project_with_tests)
def test_test_command_runs_pytest():
    p = project_with_tests()
    runner = CliRunner()
    r = runner.invoke(main, ['test', '--project', str(p)])
    assert r.exit_code == 0, r.output
    assert 'test_passes' in r.output or 'passed' in r.output


@use(project_with_tests)
def test_test_command_propagates_pytest_failure():
    p = project_with_tests()
    (p / 'tests' / 'test_fail.py').write_text(
        'def test_fails():\n    assert False\n'
    )
    runner = CliRunner()
    r = runner.invoke(main, ['test', '--project', str(p)])
    assert r.exit_code != 0


@use(project_with_tests)
def test_test_command_passes_through_pytest_args():
    p = project_with_tests()
    (p / 'tests' / 'test_other.py').write_text(
        'def test_other():\n    assert False\n'
    )
    runner = CliRunner()
    r = runner.invoke(
        main,
        ['test', '--project', str(p), '-k', 'passes'],
    )
    assert r.exit_code == 0, r.output
```

- [ ] **Step 5: Run; confirm failures**

- [ ] **Step 6: Implement**

`src/claudesheets/commands/test_cmd.py`:

```python
"""Implementation of `claudesheets test`."""

from __future__ import annotations

import subprocess
import sys
from typing import Sequence

import click

from claudesheets.exceptions import ProjectError
from claudesheets.project import Project


def run(*, project_path: str, pytest_args: Sequence[str]) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    if not project.tests_dir.is_dir():
        raise click.ClickException(
            f'no tests/ directory in {project.root}'
        )

    cmd = [sys.executable, '-m', 'pytest', str(project.tests_dir),
           *pytest_args]
    proc = subprocess.run(cmd, cwd=project.root)
    if proc.returncode != 0:
        raise click.exceptions.Exit(proc.returncode)
```

- [ ] **Step 7: Run; confirm pass**

```bash
uv run pytest tests/test_test_cmd.py -v
```

- [ ] **Step 8: Commit**

```bash
git add src/claudesheets/commands/test_cmd.py src/claudesheets/cli.py \
        src/claudesheets/project.py src/claudesheets/commands/init_cmd.py \
        tests/test_test_cmd.py
git commit -m "test: shell out to pytest for project tests"
```

---

## Task 8: Snapshot dump/load/diff helpers

**Files:**
- Create: `src/claudesheets/snapshot.py`
- Modify: `src/claudesheets/project.py` (add `snapshots_dir`)
- Create: `tests/test_snapshot.py`

- [ ] **Step 1: Add `snapshots_dir` to `Project`**

```python
    @property
    def snapshots_dir(self) -> Path:
        return self.tests_dir / 'snapshots'
```

- [ ] **Step 2: Write the failing test**

`tests/test_snapshot.py`:

```python
from claudesheets.snapshot import (
    Snapshot,
    diff_snapshots,
    snapshot_from_calc_result,
)


def test_snapshot_from_calc_result():
    cr = {'S1': {'A1': 1, 'B2': 'x'}, 'S2': {'C3': 3.5}}
    snap = snapshot_from_calc_result(cr)
    assert snap.values['S1']['A1'] == 1
    assert snap.values['S2']['C3'] == 3.5


def test_snapshot_round_trips_json(tmp_path):
    cr = {'S1': {'A1': 1}}
    snap = snapshot_from_calc_result(cr)
    p = tmp_path / 'snap.json'
    snap.write(p)
    loaded = Snapshot.read(p)
    assert loaded == snap


def test_diff_detects_changed_value():
    a = snapshot_from_calc_result({'S': {'A1': 1}})
    b = snapshot_from_calc_result({'S': {'A1': 2}})
    diffs = diff_snapshots(a, b)
    assert diffs == [('S', 'A1', 1, 2)]


def test_diff_detects_added_and_removed():
    a = snapshot_from_calc_result({'S': {'A1': 1}})
    b = snapshot_from_calc_result({'S': {'A1': 1, 'B1': 2}})
    diffs = diff_snapshots(a, b)
    assert ('S', 'B1', None, 2) in diffs


def test_no_diff_when_equal():
    a = snapshot_from_calc_result({'S': {'A1': 1}})
    b = snapshot_from_calc_result({'S': {'A1': 1}})
    assert diff_snapshots(a, b) == []
```

- [ ] **Step 3: Run; confirm failures**

- [ ] **Step 4: Implement**

`src/claudesheets/snapshot.py`:

```python
"""Golden-file snapshots of calculated workbook values."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Tuple

from claudesheets.calc.base import CalcResult
from claudesheets.model.cell import CellValue


@dataclass(frozen=True)
class Snapshot:
    values: Dict[str, Dict[str, CellValue]] = field(default_factory=dict)

    def write(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(self.values, indent=2, sort_keys=True)
        )

    @classmethod
    def read(cls, path: Path) -> 'Snapshot':
        return cls(values=json.loads(path.read_text()))


def snapshot_from_calc_result(result: CalcResult) -> Snapshot:
    return Snapshot(
        values={s: dict(cells) for s, cells in result.items()}
    )


Diff = Tuple[str, str, CellValue, CellValue]


def diff_snapshots(a: Snapshot, b: Snapshot) -> List[Diff]:
    """Return (sheet, addr, old, new) tuples for every difference.

    `None` is used in either slot to signal "missing on that side".
    """
    diffs: List[Diff] = []
    sheets = sorted(set(a.values) | set(b.values))
    for s in sheets:
        addrs = sorted(set(a.values.get(s, {})) | set(b.values.get(s, {})))
        for addr in addrs:
            va = a.values.get(s, {}).get(addr)
            vb = b.values.get(s, {}).get(addr)
            if va != vb:
                diffs.append((s, addr, va, vb))
    return diffs
```

- [ ] **Step 5: Run; confirm pass**

- [ ] **Step 6: Commit**

```bash
git add src/claudesheets/snapshot.py src/claudesheets/project.py \
        tests/test_snapshot.py
git commit -m "snapshot: dump/load/diff calculated workbook values"
```

---

## Task 9: `claudesheets snapshot` command

**Files:**
- Create: `src/claudesheets/commands/snapshot_cmd.py`
- Modify: `src/claudesheets/cli.py`
- Create: `tests/test_snapshot_cmd.py`

Snapshot files live at
`<project>/tests/snapshots/<workbook-name>.json`. When no snapshot
exists, the first run writes one (regardless of `--update`) and
reports it as "initialized".

- [ ] **Step 1: Wire CLI**

```python
@main.command('snapshot')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
)
@click.option(
    '--update',
    is_flag=True,
    help='Overwrite the saved snapshot with the current calculated values.',
)
def snapshot_cmd(project_path: str, update: bool) -> None:
    """Compare or update the golden-file snapshot of calculated values."""
    from claudesheets.commands.snapshot_cmd import run

    run(project_path=project_path, update=update)
```

- [ ] **Step 2: Write the failing test**

`tests/test_snapshot_cmd.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.testing.fixtures import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def built(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text(
        '[workbook]\nname = "in"\nsheets = []\n'
    )
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    (p / 'tests').mkdir()
    runner = CliRunner()
    runner.invoke(main, ['import', str(src), '--project', str(p)])
    runner.invoke(main, ['build', '--project', str(p)])
    yield p


@use(built, requires_libreoffice)
def test_snapshot_initializes_when_missing():
    p = built()
    runner = CliRunner()
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code == 0
    assert (p / 'tests' / 'snapshots' / 'in.json').is_file()
    assert 'initialized' in r.output.lower()


@use(built, requires_libreoffice)
def test_snapshot_clean_when_unchanged():
    p = built()
    runner = CliRunner()
    runner.invoke(main, ['snapshot', '--project', str(p)])  # init
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code == 0
    assert 'no changes' in r.output.lower()


@use(built, requires_libreoffice)
def test_snapshot_reports_diff_and_exits_nonzero():
    p = built()
    runner = CliRunner()
    runner.invoke(main, ['snapshot', '--project', str(p)])  # init
    snap = p / 'tests' / 'snapshots' / 'in.json'
    snap.write_text(snap.read_text().replace('1040000', '999999'))
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code != 0
    assert '999999' in r.output or 'differ' in r.output.lower()


@use(built, requires_libreoffice)
def test_snapshot_update_overwrites():
    p = built()
    runner = CliRunner()
    runner.invoke(main, ['snapshot', '--project', str(p)])  # init
    snap = p / 'tests' / 'snapshots' / 'in.json'
    snap.write_text('{}')
    r = runner.invoke(
        main, ['snapshot', '--project', str(p), '--update']
    )
    assert r.exit_code == 0
    text = snap.read_text()
    assert '"Outputs"' in text
```

- [ ] **Step 3: Run; confirm failures**

- [ ] **Step 4: Implement**

`src/claudesheets/commands/snapshot_cmd.py`:

```python
"""Implementation of `claudesheets snapshot`."""

from __future__ import annotations

import click

from claudesheets.calc import get_calc_engine
from claudesheets.calc.cache import (
    hash_xlsx,
    read_cached,
    write_cached,
)
from claudesheets.config import load_project
from claudesheets.exceptions import ProjectError
from claudesheets.project import Project
from claudesheets.snapshot import (
    Snapshot,
    diff_snapshots,
    snapshot_from_calc_result,
)


def run(*, project_path: str, update: bool) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    cfg = load_project(project.claudesheets_toml.read_text())
    built = project.build_dir / f'{cfg.name}.xlsx'
    if not built.is_file():
        raise click.ClickException(
            f'No built xlsx at {built}. Run `claudesheets build` first.'
        )

    key = hash_xlsx(built)
    cached = read_cached(project.calc_cache_dir, key)
    if cached is None:
        cached = get_calc_engine(cfg.calc_engine).evaluate(built)
        write_cached(project.calc_cache_dir, key, cached)

    current = snapshot_from_calc_result(cached)
    snap_path = project.snapshots_dir / f'{cfg.name}.json'

    if not snap_path.is_file():
        current.write(snap_path)
        click.echo(f'initialized snapshot at {snap_path}')
        return

    if update:
        current.write(snap_path)
        click.echo(f'updated snapshot at {snap_path}')
        return

    saved = Snapshot.read(snap_path)
    diffs = diff_snapshots(saved, current)
    if not diffs:
        click.echo('no changes')
        return

    for sheet, addr, old, new in diffs:
        click.echo(f'  {sheet}!{addr}: {old!r} -> {new!r}')
    raise click.exceptions.Exit(1)
```

- [ ] **Step 5: Run; confirm pass**

```bash
uv run pytest tests/test_snapshot_cmd.py -v
uv run mypy src/
```

- [ ] **Step 6: Commit**

```bash
git add src/claudesheets/commands/snapshot_cmd.py src/claudesheets/cli.py \
        tests/test_snapshot_cmd.py
git commit -m "snapshot: golden-file regression command"
```

---

## Task 10: End-to-end recalc/test/snapshot test

**Files:**
- Create: `tests/test_end_to_end_recalc.py`

This task wires Tasks 4–9 together against a single project to catch
integration bugs early.

- [ ] **Step 1: Write the test**

`tests/test_end_to_end_recalc.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.testing import Model
from claudesheets.testing.fixtures import requires_libreoffice
from tests.fixtures.workbooks import write_simple_xlsx


@fixture
def project(tmp_path: Path):
    src = tmp_path / 'in.xlsx'
    write_simple_xlsx(src)
    p = tmp_path / 'proj'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "in"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text(
        '[workbook]\nname = "in"\nsheets = []\n'
    )
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    (p / 'tests').mkdir()
    runner = CliRunner()
    runner.invoke(main, ['import', str(src), '--project', str(p)])
    runner.invoke(main, ['build', '--project', str(p)])
    yield p


@use(project, requires_libreoffice)
def test_full_recalc_then_snapshot_then_model_set():
    p = project()
    runner = CliRunner()

    r = runner.invoke(main, ['recalc', '--project', str(p)])
    assert r.exit_code == 0, r.output

    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code == 0, r.output

    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code == 0, r.output
    assert 'no changes' in r.output.lower()

    m = Model.open(p)
    assert m.get('Outputs!B1') == 1_040_000
    m.set('Inputs!B1', 0.10)
    assert m.get('Outputs!B1') == 1_100_000
```

- [ ] **Step 2: Run; expect pass**

```bash
uv run pytest tests/test_end_to_end_recalc.py -v
```

- [ ] **Step 3: Run the full suite**

```bash
uv run pytest -v
```

Expected: all green (skips for libreoffice tests are acceptable on
machines without `soffice`).

- [ ] **Step 4: Commit**

```bash
git add tests/test_end_to_end_recalc.py
git commit -m "tests: end-to-end recalc/snapshot/model"
```

---

## Task 11: Documentation pass

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update Status**

Replace the current Status section with:

```
## Status

Plan 2 complete: a swappable calc engine (LibreOffice headless),
content-addressed calc cache, and `recalc`, `test`, and `snapshot`
commands. The `claudesheets.testing.Model` API gives pytest tests
`set/get/recalc` over a built workbook. Diff/check, conditional
formatting, comments, and the escape-hatch re-import flow are coming
in Plans 3–4.
```

- [ ] **Step 2: Add a brief Calc engine section**

Append after the Quick reference:

```
## Calc engine

The default engine is LibreOffice headless. `soffice` must be on
`$PATH` (Debian/Ubuntu: `apt install libreoffice`; macOS:
`brew install --cask libreoffice`).

The engine is selected per-project in `claudesheets.toml`:

\`\`\`toml
[build]
calc_engine = "libreoffice"
\`\`\`

The interface is documented in `src/claudesheets/calc/base.py`;
implement `CalcEngine.evaluate` to add a new backend.
```

- [ ] **Step 3: Verify**

```bash
uv run pytest -q
uv run claudesheets --help
```

- [ ] **Step 4: Commit**

```bash
git add README.md
git commit -m "docs: update status after Plan 2"
```

---

## Done criteria for Plan 2

1. `uv run pytest -v` is green (libreoffice-gated tests skip cleanly
   when `soffice` is missing; on machines with it, they pass).
2. `claudesheets recalc` produces a cache entry; second run reports
   `cache hit`.
3. `claudesheets test` runs the project's pytest tests and propagates
   exit codes.
4. `claudesheets snapshot` initializes, reports clean, reports diffs,
   and supports `--update`.
5. `from claudesheets.testing import Model` followed by
   `Model.open(project).get('Sheet!A1')` returns the calculated
   value, and `.set(addr, v)` invalidates and recomputes.

## What this plan does **not** deliver (deferred)

- `diff` and `check` commands — Plan 4
- Conditional formatting, cell comments, frozen panes, print areas,
  ListObject tables — Plan 3
- Hash-based detection of external xlsx edits + interactive merge
  re-import — Plan 4
- `--flatten` for external references — Plan 4
- MCP wrapper — Plan 5
- Alternative calc engines (xlwings, Microsoft Graph) — beyond v1
