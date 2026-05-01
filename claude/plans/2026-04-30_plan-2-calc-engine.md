# Plan 2 — Calc engine, testing library, recalc/test/snapshot

> **For agentic workers:** REQUIRED SUB-SKILL: Use
> `superpowers:subagent-driven-development` (recommended) or
> `superpowers:executing-plans` to implement this plan task-by-task.
> Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Wire up a swappable calc engine with a LibreOffice headless
implementation, expose a `claudesheets.testing.Model` API for pytest
tests, and ship the `recalc`, `test`, and `snapshot` commands.

**Architecture:**
- The `xlsx` writer is made deterministic by pinning
  `core_properties.created`/`modified` so the same source produces
  byte-identical output. This makes the calc cache key meaningful.
- A `CalcEngine` ABC defines `evaluate(xlsx_path) -> dict[sheet, dict[addr, value]]`.
- `LibreOfficeEngine` implements it via `soffice --headless --calc
  --convert-to xlsx` to force a recalc, then reads cached values back
  with `openpyxl(..., data_only=True)`.
- A content-addressed calc cache (`.claudesheets/calc/<sha256>.json`)
  avoids re-running the engine when the built xlsx hasn't changed.
- `claudesheets.testing.Model` wraps a `Workbook` + `CalcEngine` and
  gives user tests a `Model.set/get/recalc` surface.
- `recalc` builds the xlsx, runs the engine, writes the cache.
- `test` invokes [testsweet](https://github.com/kaapstorm/testsweet)
  in-process against `<project>/tests/`. No subprocess, no
  pytest-config dependency.
- `snapshot` reads the cache (running `recalc` first if missing),
  filters the result down to formula cells only (the spec's
  "calculated outputs"), normalizes datetimes to ISO strings, and
  diffs/updates `tests/snapshots/<workbook>.json`.

**Tech stack:** Python 3.11, openpyxl, click, subprocess, hashlib,
LibreOffice (`soffice` on `$PATH`), testsweet (runtime dep, used by
`claudesheets test`), pytest + pytest-unmagic (dev deps, only for
claudesheets's own internal tests).

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
├── __init__.py            # public API: Model, parse_address, require_libreoffice
├── addresses.py           # parse_address helper
├── model.py               # Model class
└── runtime.py             # require_libreoffice() helper for user tests

src/claudesheets/commands/
├── recalc_cmd.py          # implementation of `recalc`
├── test_cmd.py            # implementation of `test`
└── snapshot_cmd.py        # implementation of `snapshot`

src/claudesheets/snapshot.py  # snapshot dump/load/diff helpers

tests/
├── fixtures/
│   └── libreoffice.py         # internal-only `requires_libreoffice` pytest fixture
├── test_calc_libreoffice.py   # gated on `soffice` being available
├── test_calc_cache.py
├── test_xlsx_determinism.py   # `build` produces byte-identical xlsx
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
- `src/claudesheets/xlsx/writer.py` — pin `core_properties` for
  deterministic builds (Task 3)
- `pyproject.toml` — add `testsweet` as a runtime dependency (Task 0)
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

## Task 0: Scaffold packages and add the testsweet runtime dependency

**Files:**
- Create: `src/claudesheets/calc/__init__.py`
- Create: `src/claudesheets/testing/__init__.py`
- Modify: `pyproject.toml` (via `uv add testsweet`)

- [ ] **Step 1: Add testsweet as a runtime dependency**

```bash
uv add 'testsweet>=0.1.4'
```

This must be a runtime dep (not dev-only) because `claudesheets test`
imports `testsweet.__main__:main` to drive user-project tests
in-process.

- [ ] **Step 2: Create the empty packages**

`src/claudesheets/calc/__init__.py`:

```python
"""Calc engine plugin layer.

A calc engine evaluates a built `.xlsx` and returns the calculated
values for every cell. The default engine is LibreOffice headless.
"""
```

`src/claudesheets/testing/__init__.py`:

```python
"""Test-time helpers exposed to user testsweet tests."""
```

- [ ] **Step 3: Verify**

```bash
uv run python -c "import claudesheets.calc, claudesheets.testing, testsweet"
```

- [ ] **Step 4: Commit**

```bash
git add pyproject.toml uv.lock \
        src/claudesheets/calc/__init__.py \
        src/claudesheets/testing/__init__.py
git commit -m "calc: scaffold calc + testing packages, add testsweet dep"
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
- Create: `tests/fixtures/libreoffice.py`
- Create: `tests/test_calc_libreoffice.py`

The engine recalculates by re-saving the workbook with LibreOffice,
which forces formula evaluation, then reads `data_only=True` to
extract calculated values.

The `requires_libreoffice` skip fixture lives under `tests/fixtures/`
because it's only useful to claudesheets's *internal* pytest suite.
User-facing tests run under testsweet (see Task 8) and use a
different mechanism — the public `claudesheets.testing.runtime`
helpers.

- [ ] **Step 1: Write the internal-only `requires_libreoffice` fixture**

`tests/fixtures/libreoffice.py`:

```python
"""Internal-only pytest fixture: skip when LibreOffice is missing.

Lives under tests/ (not the public package) because it depends on
pytest, while the user-facing test framework is testsweet.
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

import pytest
from unmagic import use

from claudesheets.calc.libreoffice import LibreOfficeEngine, LibreOfficeError
from tests.fixtures.libreoffice import requires_libreoffice
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


def test_libreoffice_engine_raises_when_binary_missing(tmp_path: Path):
    # No need for soffice on $PATH — we point at a path that doesn't exist.
    src = tmp_path / 'fake.xlsx'
    src.write_bytes(b'not really an xlsx')
    eng = LibreOfficeEngine(soffice='/no/such/soffice/binary')
    with pytest.raises(LibreOfficeError, match='not on'):
        eng.evaluate(src)
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
        tests/fixtures/libreoffice.py \
        tests/test_calc_libreoffice.py
git commit -m "calc: LibreOffice headless engine"
```

---

## Task 3: Deterministic xlsx writes

**Files:**
- Modify: `src/claudesheets/xlsx/writer.py`
- Create: `tests/test_xlsx_determinism.py`

`openpyxl` writes `core_properties.created` and `.modified`
timestamps to the current wall-clock time on every save, which means
the same source produces different xlsx bytes on consecutive builds.
The calc cache (Task 4) keys on the SHA-256 of the xlsx, so without
this fix it never hits in practice. Pinning the timestamps to a fixed
epoch makes builds reproducible — useful both for the cache and for
diffing `imports/*.xlsx` later.

- [ ] **Step 1: Write the failing test**

`tests/test_xlsx_determinism.py`:

```python
import hashlib
import time
from pathlib import Path

from unmagic import use

from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


def _sha256(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def test_two_builds_of_same_workbook_are_byte_identical(tmp_path: Path):
    from claudesheets.xlsx.reader import read_xlsx

    src = tmp_path / 'src.xlsx'
    write_simple_xlsx(src)
    wb = read_xlsx(src)

    a = tmp_path / 'a.xlsx'
    b = tmp_path / 'b.xlsx'
    write_xlsx(wb, a)
    time.sleep(1.1)  # ensure wall-clock would differ
    write_xlsx(wb, b)

    assert _sha256(a) == _sha256(b), (
        'xlsx writer is not deterministic: same Workbook produced '
        'different bytes 1s apart'
    )
```

- [ ] **Step 2: Run; confirm failure**

```bash
uv run pytest tests/test_xlsx_determinism.py -v
```

Expected: failure (the existing writer embeds wall-clock timestamps).

- [ ] **Step 3: Pin timestamps in the writer**

In `src/claudesheets/xlsx/writer.py`, the openpyxl `Workbook` is the
local variable `out` and the save call is `out.save(path)` near the
end of the function. Add a module-level epoch and pin the properties
just before the save:

```python
# Near the top of the module, with other imports:
from datetime import datetime

# Module-level constant:
_DETERMINISTIC_EPOCH = datetime(2000, 1, 1)

# Inside write_xlsx, immediately before `out.save(path)`:
    out.properties.created = _DETERMINISTIC_EPOCH
    out.properties.modified = _DETERMINISTIC_EPOCH
    out.save(path)
```

Why a module-level constant: keeps the value in one place if a future
change wants to bump it (or expose it for debugging).

- [ ] **Step 4: Run; confirm pass**

```bash
uv run pytest tests/test_xlsx_determinism.py -v
uv run pytest tests/test_xlsx_roundtrip.py -v   # regression check
uv run mypy src/
```

- [ ] **Step 5: Commit**

```bash
git add src/claudesheets/xlsx/writer.py tests/test_xlsx_determinism.py
git commit -m "xlsx: deterministic builds (pin core_properties)"
```

---

## Task 4: Calc cache

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

## Task 5: `claudesheets recalc` command

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
from tests.fixtures.libreoffice import requires_libreoffice
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
    from claudesheets.commands.recalc_cmd import CACHE_HIT_MESSAGE

    p = imported()
    runner = CliRunner()
    runner.invoke(main, ['recalc', '--project', str(p)])
    out = runner.invoke(main, ['recalc', '--project', str(p)])
    assert out.exit_code == 0
    assert CACHE_HIT_MESSAGE in out.output


@use(imported, requires_libreoffice)
def test_recalc_force_rebuilds_cache():
    from claudesheets.commands.recalc_cmd import CACHE_HIT_MESSAGE

    p = imported()
    runner = CliRunner()
    runner.invoke(main, ['recalc', '--project', str(p)])
    out = runner.invoke(
        main, ['recalc', '--project', str(p), '--force']
    )
    assert out.exit_code == 0
    assert CACHE_HIT_MESSAGE not in out.output
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


CACHE_HIT_MESSAGE = 'cache hit'


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
            click.echo(f'{CACHE_HIT_MESSAGE}: {key[:12]}')
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

## Task 6: Address parsing helpers

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


def test_parse_sheet_scoped_named_range_resolves():
    wb = Workbook(name='x', sheets=[Sheet(name='Inputs')])
    wb.named_ranges.append(
        NamedRange(
            name='local_rate',
            ref='Inputs!$B$5',
            scope='sheet',
            sheet='Inputs',
        )
    )
    # Sheet-scoped names resolve identically to workbook-scoped names
    # for our purposes: parse_address looks up by name and follows the
    # `ref`. Disambiguation between two sheet-scoped names with the
    # same string is out of scope for Plan 2 (no real workbook does
    # that).
    assert parse_address(wb, 'local_rate') == ('Inputs', 'B5')
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

## Task 7: `claudesheets.testing.Model`

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
from tests.fixtures.libreoffice import requires_libreoffice
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

## Task 8: `claudesheets test` command (testsweet, in-process)

**Files:**
- Create: `src/claudesheets/commands/test_cmd.py`
- Modify: `src/claudesheets/cli.py`
- Modify: `src/claudesheets/project.py` (add `tests_dir`)
- Modify: `src/claudesheets/commands/init_cmd.py` (scaffold `tests/`)
- Modify: `tests/test_init.py` (assert tests/ scaffolded)
- Create: `tests/test_test_cmd.py`

User-project tests run under [testsweet](https://github.com/kaapstorm/testsweet),
not pytest. Tests are decorated with `@test`; the runner is invoked
**in-process** via `testsweet.__main__.main(argv) -> int`. This
sidesteps the entire `sys.executable`/subprocess discovery problem
because tests execute in the same Python that runs the CLI — the
user simply needs to install claudesheets into their project venv
(documented in Task 12).

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

Extend `tests/test_init.py::test_init_creates_project_skeleton` to
also assert:

```python
    assert (project / 'tests').is_dir()
    assert (project / 'tests' / '__init__.py').is_file()
```

- [ ] **Step 3: Wire CLI**

In `src/claudesheets/cli.py`:

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
@click.argument('targets', nargs=-1, type=click.UNPROCESSED)
def test_cmd(project_path: str, targets: tuple[str, ...]) -> None:
    """Run the project's testsweet tests."""
    from claudesheets.commands.test_cmd import run

    run(project_path=project_path, targets=list(targets))
```

- [ ] **Step 4: Write the failing test**

`tests/test_test_cmd.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main


_PASS_TEST = """\
from testsweet import test


@test
def passes():
    assert 1 + 1 == 2
"""

_FAIL_TEST = """\
from testsweet import test


@test
def fails():
    assert False
"""


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
    (p / 'tests' / 'test_simple.py').write_text(_PASS_TEST)
    yield p


@use(project_with_tests)
def test_test_command_runs_testsweet():
    p = project_with_tests()
    runner = CliRunner()
    r = runner.invoke(main, ['test', '--project', str(p)])
    assert r.exit_code == 0, r.output
    assert 'ok' in r.output
    assert 'passes' in r.output


@use(project_with_tests)
def test_test_command_propagates_failure_exit_code():
    p = project_with_tests()
    (p / 'tests' / 'test_fail.py').write_text(_FAIL_TEST)
    runner = CliRunner()
    r = runner.invoke(main, ['test', '--project', str(p)])
    assert r.exit_code != 0
    assert 'FAIL' in r.output


@use(project_with_tests)
def test_test_command_supports_target_selection():
    p = project_with_tests()
    (p / 'tests' / 'test_other.py').write_text(_FAIL_TEST)
    runner = CliRunner()
    # Pass a specific target so we only run the passing test.
    r = runner.invoke(
        main,
        ['test', '--project', str(p), 'tests/test_simple.py'],
    )
    assert r.exit_code == 0, r.output


@use(project_with_tests)
def test_test_command_errors_when_tests_dir_missing(tmp_path: Path):
    p = tmp_path / 'no-tests'
    p.mkdir()
    (p / 'claudesheets.toml').write_text(
        '[project]\nname = "x"\n[build]\ncalc_engine = "libreoffice"\n'
    )
    (p / 'workbook.toml').write_text(
        '[workbook]\nname = "x"\nsheets = []\n'
    )
    (p / 'sheets').mkdir()
    (p / 'data').mkdir()
    runner = CliRunner()
    r = runner.invoke(main, ['test', '--project', str(p)])
    assert r.exit_code != 0
    assert 'tests/' in r.output or 'tests' in r.output.lower()
```

- [ ] **Step 5: Run; confirm failures**

- [ ] **Step 6: Implement**

`src/claudesheets/commands/test_cmd.py`:

```python
"""Implementation of `claudesheets test` (testsweet, in-process).

We invoke testsweet's `main(argv) -> int` programmatically. testsweet
saves and restores `sys.path` itself; we additionally save/restore
`os.getcwd()` because testsweet reads `[tool.testsweet.discovery]`
config from the cwd's `pyproject.toml`.
"""

from __future__ import annotations

import os
from typing import Sequence

import click

from claudesheets.exceptions import ProjectError
from claudesheets.project import Project


def run(*, project_path: str, targets: Sequence[str]) -> None:
    try:
        project = Project.open(project_path)
    except ProjectError as e:
        raise click.ClickException(str(e))

    if not project.tests_dir.is_dir():
        raise click.ClickException(
            f'no tests/ directory in {project.root}'
        )

    from testsweet.__main__ import main as testsweet_main

    argv = list(targets) if targets else [str(project.tests_dir)]
    prev_cwd = os.getcwd()
    try:
        os.chdir(project.root)
        rc = testsweet_main(argv)
    finally:
        os.chdir(prev_cwd)

    if rc != 0:
        raise click.exceptions.Exit(rc)
```

- [ ] **Step 7: Run; confirm pass**

```bash
uv run pytest tests/test_test_cmd.py -v
```

- [ ] **Step 8: Commit**

```bash
git add src/claudesheets/commands/test_cmd.py src/claudesheets/cli.py \
        src/claudesheets/project.py src/claudesheets/commands/init_cmd.py \
        tests/test_test_cmd.py tests/test_init.py
git commit -m "test: run user testsweet tests in-process"
```

---

## Task 9: Snapshot dump/load/diff helpers

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

Snapshots include **only formula cells** — the spec's "calculated
outputs". Literal inputs are excluded so changing an input doesn't
generate noise at the input cell itself; the diff highlights only
its propagation downstream.

`datetime` cell values are normalized to ISO-8601 strings before the
in-memory `Snapshot` is constructed, so JSON write/read is lossless
for equality purposes.

- [ ] **Step 2: Write the failing test**

`tests/test_snapshot.py`:

```python
from datetime import datetime

from claudesheets.model.cell import Cell
from claudesheets.model.workbook import Sheet, Workbook
from claudesheets.snapshot import (
    Snapshot,
    diff_snapshots,
    snapshot_from_calc_result,
)


def _wb_with_formulas(formula_addrs):
    """A workbook where the named addresses are formula cells."""
    sheets_by_name = {}
    for sheet_name, addr in formula_addrs:
        s = sheets_by_name.setdefault(sheet_name, Sheet(name=sheet_name))
        s.set(addr, Cell(formula='=1+1'))
    wb = Workbook(name='x', sheets=list(sheets_by_name.values()))
    return wb


def test_snapshot_includes_only_formula_cells():
    # Source: A1 is literal, B1 is a formula. CalcResult has both.
    wb = _wb_with_formulas([('S1', 'B1')])
    cr = {'S1': {'A1': 'literal', 'B1': 42}}
    snap = snapshot_from_calc_result(cr, wb)
    assert snap.values == {'S1': {'B1': 42}}


def test_snapshot_normalizes_datetime_to_iso_string():
    wb = _wb_with_formulas([('S1', 'A1')])
    cr = {'S1': {'A1': datetime(2024, 3, 15, 12, 0, 0)}}
    snap = snapshot_from_calc_result(cr, wb)
    assert snap.values['S1']['A1'] == '2024-03-15T12:00:00'


def test_snapshot_round_trips_json(tmp_path):
    wb = _wb_with_formulas([('S1', 'A1')])
    snap = snapshot_from_calc_result({'S1': {'A1': 1}}, wb)
    p = tmp_path / 'snap.json'
    snap.write(p)
    loaded = Snapshot.read(p)
    assert loaded == snap


def test_diff_detects_changed_value():
    wb = _wb_with_formulas([('S', 'A1')])
    a = snapshot_from_calc_result({'S': {'A1': 1}}, wb)
    b = snapshot_from_calc_result({'S': {'A1': 2}}, wb)
    diffs = diff_snapshots(a, b)
    assert diffs == [('S', 'A1', 1, 2)]


def test_diff_detects_added_and_removed():
    wb_a = _wb_with_formulas([('S', 'A1')])
    wb_b = _wb_with_formulas([('S', 'A1'), ('S', 'B1')])
    a = snapshot_from_calc_result({'S': {'A1': 1}}, wb_a)
    b = snapshot_from_calc_result({'S': {'A1': 1, 'B1': 2}}, wb_b)
    diffs = diff_snapshots(a, b)
    assert ('S', 'B1', None, 2) in diffs


def test_no_diff_when_equal():
    wb = _wb_with_formulas([('S', 'A1')])
    a = snapshot_from_calc_result({'S': {'A1': 1}}, wb)
    b = snapshot_from_calc_result({'S': {'A1': 1}}, wb)
    assert diff_snapshots(a, b) == []
```

- [ ] **Step 3: Run; confirm failures**

- [ ] **Step 4: Implement**

`src/claudesheets/snapshot.py`:

```python
"""Golden-file snapshots of calculated workbook values.

A `Snapshot` captures the calculated value of every formula cell in
a workbook. Literal-input cells are intentionally excluded.

Datetime values are normalized to ISO-8601 strings before snapshot
construction so JSON round-trips preserve equality.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

from claudesheets.calc.base import CalcResult
from claudesheets.model.cell import CellValue
from claudesheets.model.workbook import Workbook


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


def snapshot_from_calc_result(
    result: CalcResult, workbook: Workbook
) -> Snapshot:
    """Build a snapshot from a CalcResult, keeping only formula cells.

    `workbook` is consulted to determine which cells in the result
    were formulas in the source. Literal inputs are dropped.
    """
    formulas: Dict[str, set[str]] = {}
    for sheet in workbook.sheets:
        formulas[sheet.name] = {
            addr for addr, cell in sheet.cells.items()
            if cell.formula is not None
        }

    values: Dict[str, Dict[str, CellValue]] = {}
    for sheet_name, cells in result.items():
        keep = formulas.get(sheet_name, set())
        sheet_out: Dict[str, CellValue] = {}
        for addr, value in cells.items():
            if addr not in keep:
                continue
            sheet_out[addr] = _normalize(value)
        if sheet_out:
            values[sheet_name] = sheet_out
    return Snapshot(values=values)


def _normalize(value: CellValue) -> CellValue:
    if isinstance(value, datetime):
        return value.isoformat()
    return value


Diff = Tuple[str, str, CellValue, CellValue]


def diff_snapshots(a: Snapshot, b: Snapshot) -> List[Diff]:
    """Return (sheet, addr, old, new) tuples for every difference.

    `None` is used in either slot to signal "missing on that side".
    """
    diffs: List[Diff] = []
    sheets = sorted(set(a.values) | set(b.values))
    for s in sheets:
        addrs = sorted(
            set(a.values.get(s, {})) | set(b.values.get(s, {}))
        )
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

## Task 10: `claudesheets snapshot` command

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
from tests.fixtures.libreoffice import requires_libreoffice
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
    import json

    p = built()
    runner = CliRunner()
    runner.invoke(main, ['snapshot', '--project', str(p)])  # init
    snap = p / 'tests' / 'snapshots' / 'in.json'
    data = json.loads(snap.read_text())
    # Mutate the formula cell Outputs!B1 so the next snapshot diffs.
    data['Outputs']['B1'] = 999_999
    snap.write_text(json.dumps(data, indent=2, sort_keys=True))
    r = runner.invoke(main, ['snapshot', '--project', str(p)])
    assert r.exit_code != 0
    assert '999999' in r.output or 'Outputs!B1' in r.output


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
from claudesheets.source.reader import read_source


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

    workbook = read_source(project.root)
    current = snapshot_from_calc_result(cached, workbook)
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

## Task 11: End-to-end recalc/test/snapshot test

**Files:**
- Create: `tests/test_end_to_end_recalc.py`

This task wires Tasks 5–10 together against a single project to catch
integration bugs early.

- [ ] **Step 1: Write the test**

`tests/test_end_to_end_recalc.py`:

```python
from pathlib import Path

from click.testing import CliRunner
from unmagic import fixture, use

from claudesheets.cli import main
from claudesheets.testing import Model
from tests.fixtures.libreoffice import requires_libreoffice
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

## Task 12: Documentation pass

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update Status**

Replace the current Status section with:

```
## Status

Plan 2 complete: a swappable calc engine (LibreOffice headless),
deterministic xlsx builds, content-addressed calc cache, and
`recalc`, `test`, and `snapshot` commands. The
`claudesheets.testing.Model` API gives [testsweet](https://github.com/kaapstorm/testsweet)
tests `set/get/recalc` over a built workbook. Diff/check, conditional
formatting, comments, and the escape-hatch re-import flow are coming
in Plans 3–4.
```

- [ ] **Step 2: Add Calc engine and Testing sections**

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

## Testing your model

Tests use [testsweet](https://github.com/kaapstorm/testsweet) — plain
Python functions decorated with `@test`. Install claudesheets into
your project venv (not via `uv tool install`, which isolates
claudesheets from your project's dependencies):

\`\`\`bash
uv add claudesheets
# or, if not using uv:
pip install claudesheets
\`\`\`

Then write tests under `tests/`:

\`\`\`python
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
\`\`\`

Run them with `claudesheets test` (in-process testsweet) or directly
with `python -m testsweet tests/`.
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
3. `claudesheets test` runs the project's testsweet tests in-process
   and propagates exit codes.
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
