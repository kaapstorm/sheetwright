# Testing reference

The `claudesheets.testing` API plus testsweet patterns for asserting
on workbook outputs.

## `Model`

`claudesheets.testing.Model` wraps a workbook plus a calc engine.
`set` mutates the in-memory workbook and invalidates calculated
values; `get` triggers a recalc on demand and returns the
calculated value (or the literal value, for cells that aren't
formulas).

### `Model.open(project_path)`

Load a project from disk and return a fresh `Model`. The project is
read from `claudesheets.toml` and `workbook.toml`; the calc engine
is selected by `claudesheets.toml`'s `build.calc_engine`.

```python
from claudesheets.testing import Model

model = Model.open('.')
```

`project_path` may be a `str` or `pathlib.Path`. Tests typically
pass `'.'` so the project root resolves relative to wherever the
test runner is invoked.

### `Model(workbook, engine)`

Construct directly when you want a hermetic in-memory test that
doesn't read the filesystem.

```python
from claudesheets.calc.base import CalcEngine, CalcResult
from claudesheets.model.cell import Cell
from claudesheets.model.workbook import Sheet, Workbook
from claudesheets.testing import Model


class StubEngine(CalcEngine):
    def evaluate(self, xlsx_path):
        return {'S': {'A1': 42}}


wb = Workbook(name='m', sheets=[Sheet(name='S')])
wb.sheet('S').set('A1', Cell(formula='=1+41'))
model = Model(wb, StubEngine())
assert model.get('S!A1') == 42
```

This bypasses LibreOffice; useful for testing your own helpers.

### `model.set(address, value)`

Mutate a cell in the in-memory workbook and invalidate any cached
calculation.

```python
model.set('Inputs!B1', 0.05)
model.set("'Sheet With Space'!A1", 'hello')
model.set('growth_rate', 0.05)         # named range
```

Address forms:

- `Sheet!A1` — Excel-style sheet-qualified address.
- `'Sheet With Space'!A1` — quote sheet names containing spaces; an
  embedded apostrophe is escaped as `''`.
- `growth_rate` — workbook-scoped named range whose `ref` resolves
  to a single cell.
- Bare `A1` is **rejected**. Tests must say which sheet they mean.

`value` may be `str`, `int`, `float`, `bool`, `datetime`, or `None`.

### `model.get(address)`

Return the calculated value (or, for non-formula cells, the literal
value).

```python
revenue = model.get('Outputs!B1')
assert math.isclose(revenue, 1_234_567, rel_tol=1e-6)
```

`get` triggers a recalc if the workbook has been mutated since the
last calculation. If the cell holds a formula and the calc engine
omitted it from its result, `get` raises `RuntimeError` with the
sheet and address — never silently returns `None`.

### `model.recalc()`

Force a recalc explicitly. Rarely needed: `get` does this lazily.

```python
model.set('Inputs!B1', 0.05)
model.recalc()
# subsequent get() calls hit the fresh calculation
```

`recalc` writes the workbook to a temp xlsx and invokes the calc
engine. With LibreOffice that's a subprocess with a dedicated
profile (so concurrent tests don't fight over the user-profile
lock). Expect 1–3 seconds per recalc on a small model.

### `model.workbook`

The underlying `Workbook` object, for tests that want to inspect
structure (e.g. "this sheet has the expected named ranges") rather
than values.

## testsweet patterns

claudesheets uses [testsweet](https://github.com/kaapstorm/testsweet)
for tests, which differs from pytest in three ways relevant here:

- `@test` is the function decorator (no implicit "function named
  `test_*`" rule).
- Fixtures are imported and called explicitly (no parameter-name
  injection).
- `catch_exceptions` is the idiom for asserting that something
  raises.

### `@test`

```python
from testsweet import test


@test
def revenue_grows_with_assumption():
    model = Model.open('.')
    model.set('growth_rate', 0.05)
    assert math.isclose(model.get('Outputs!B1'), 1_050_000, rel_tol=1e-9)
```

### `@test_params`

Run the same test body across multiple parameter sets:

```python
from testsweet import test_params


@test_params([
    {'rate': 0.00, 'expected': 1_000_000},
    {'rate': 0.05, 'expected': 1_050_000},
    {'rate': 0.10, 'expected': 1_100_000},
])
def revenue_scales_linearly(rate, expected):
    model = Model.open('.')
    model.set('growth_rate', rate)
    assert math.isclose(model.get('Outputs!B1'), expected, rel_tol=1e-9)
```

### `catch_exceptions`

```python
from testsweet import catch_exceptions, test


@test
def negative_growth_rejected():
    model = Model.open('.')
    with catch_exceptions() as excs:
        model.set('growth_rate', -0.5)
        model.get('Outputs!B1')
    assert excs and isinstance(excs[0], ValueError)
```

### `@skip`

```python
from testsweet import skip, test


@test
@skip('pending soffice 7.6 in CI')
def conditional_format_renders():
    ...
```

### Fixtures

Define fixtures with `@unmagic.fixture`, yield once, and apply with
`@use(fixture)` or the shorthand `@<fixture>`. Don't put them in
`conftest.py` — keep them in a regular module and import them where
used.

```python
from unmagic import fixture, use
from testsweet import test

from claudesheets.testing import Model


@fixture
def model():
    m = Model.open('.')
    yield m


@test
@use(model)
def revenue_grows():
    m = model()
    m.set('growth_rate', 0.05)
    assert math.isclose(m.get('Outputs!B1'), 1_050_000, rel_tol=1e-9)
```

For expensive setup (a recalc-heavy baseline workbook), use
`scope='module'`:

```python
@fixture(scope='module')
def warm_model():
    m = Model.open('.')
    m.recalc()              # pay the LibreOffice cost once
    yield m
```

## A complete example

`tests/test_revenue.py`:

```python
import math

from testsweet import test, test_params

from claudesheets.testing import Model


@test
def baseline_revenue_at_default_growth_rate():
    model = Model.open('.')
    assert math.isclose(model.get('Outputs!B1'), 1_040_000, rel_tol=1e-9)


@test_params([
    {'rate': 0.00, 'expected': 1_000_000},
    {'rate': 0.05, 'expected': 1_050_000},
    {'rate': 0.10, 'expected': 1_100_000},
])
def revenue_scales_with_growth_rate(rate, expected):
    model = Model.open('.')
    model.set('growth_rate', rate)
    assert math.isclose(model.get('Outputs!B1'), expected, rel_tol=1e-9)


@test
def named_range_resolves_to_inputs_b1():
    model = Model.open('.')
    assert model.workbook.named_ranges[0].name == 'growth_rate'
```

Run with:

```bash
claudesheets test
```

Or, to use testsweet's own discovery:

```bash
uv run python -m testsweet tests/
```

## Performance notes

Each `model.recalc()` (and the first `model.get()` after a `set()`)
runs LibreOffice headless. That's ~1–3 seconds per call on a small
model. For test suites with hundreds of assertions, prefer
parameterised tests over many individual tests, share `Model`
instances within a test where possible, and use `scope='module'`
fixtures for read-only baselines.

## See also

- [Calc engine reference](calc-engine.md) for the underlying
  evaluation contract.
- [TDD workflow tutorial](../tutorials/tdd-workflow.md) for a worked
  example.
- [Snapshots tutorial](../tutorials/snapshots.md) when you want
  every formula cell asserted on automatically.
