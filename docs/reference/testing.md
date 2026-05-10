# Testing reference

The `sheetwright.testing` API plus testsweet patterns for asserting
on workbook outputs.

## `Model`

`sheetwright.testing.Model` wraps a workbook plus a calc engine.
`set` mutates the in-memory workbook and invalidates calculated
values; `get` triggers a recalc on demand and returns the
calculated value (or the literal value, for cells that aren't
formulas).

### `Model.open(project_path)`

Load a project from disk and return a fresh `Model`. The project is
read from `sheetwright.toml` and `workbook.toml`; the calc engine
is selected by `sheetwright.toml`'s `build.calc_engine`.

```python
from sheetwright.testing import Model

model = Model.open('.')
```

`project_path` may be a `str` or `pathlib.Path`. Tests typically
pass `'.'` so the project root resolves relative to wherever the
test runner is invoked.

### `Model(workbook, engine)`

Construct directly when you want a hermetic in-memory test that
doesn't read the filesystem.

```python
from sheetwright.calc.base import CalcEngine, CalcResult
from sheetwright.model.cell import Cell
from sheetwright.model.workbook import Sheet, Workbook
from sheetwright.testing import Model


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

sheetwright uses [testsweet](https://github.com/kaapstorm/testsweet)
for tests, which differs from pytest in a few ways relevant here:

- `@test` is an explicit decorator (no implicit "function named
  `test_*`" rule).
- There is no fixture system. Use plain context managers for
  function-style tests, or implement the context-manager protocol on
  a class for class-style tests.
- `catch_exceptions` is the idiom for asserting that something
  raises.
- `@xfail` is **strict**: an unexpected pass fails the run.

### `@test`

Mark a function as a test:

```python
from testsweet import test


@test
def revenue_grows_with_assumption():
    model = Model.open('.')
    model.set('growth_rate', 0.05)
    assert math.isclose(model.get('Outputs!B1'), 1_050_000, rel_tol=1e-9)
```

Or mark a class — every public method (not starting with `_`) is run
as a test:

```python
@test
class RevenueModel:
    def baseline(self):
        model = Model.open('.')
        assert math.isclose(
            model.get('Outputs!B1'), 1_040_000, rel_tol=1e-9,
        )

    def grows_with_assumption(self):
        model = Model.open('.')
        model.set('growth_rate', 0.05)
        assert math.isclose(
            model.get('Outputs!B1'), 1_050_000, rel_tol=1e-9,
        )
```

### `@params`

Run the same test body across multiple parameter sets. Stack with
`@test` to register the function for discovery; each tuple in the
iterable is unpacked as positional arguments. The iterable is
materialized eagerly at decoration time:

```python
from testsweet import params, test


@test
@params([
    (0.00, 1_000_000),
    (0.05, 1_050_000),
    (0.10, 1_100_000),
])
def revenue_scales_linearly(rate, expected):
    model = Model.open('.')
    model.set('growth_rate', rate)
    assert math.isclose(model.get('Outputs!B1'), expected, rel_tol=1e-9)
```

Use `@params_lazy` instead when materializing the iterable is
expensive or has side effects you want deferred until run time.

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

`catch_warnings` is the warning-capture analogue:

```python
from testsweet import catch_warnings

with catch_warnings() as warns:
    ...
assert any(isinstance(w, DeprecationWarning) for w in warns)
```

### `@skip`

`@skip` accepts `reason=` and `condition=` keyword arguments. Bare
`@skip` always skips:

```python
from testsweet import skip, test


@test
@skip(reason='pending soffice 7.6 in CI')
def conditional_format_renders():
    ...


@test
@skip(condition=sys.platform == 'win32', reason='posix-only')
def uses_named_pipe():
    ...
```

`condition=` accepts a bool or a zero-arg callable (the callable is
evaluated at run time).

### `@xfail`

Mark a test as expected to fail. If it raises, the runner reports
`xfailed`; if it unexpectedly passes, the runner reports `XPASSED`
and the run fails. Either remove the marker (the bug is fixed) or
fix the test.

```python
from testsweet import test, xfail


@test
@xfail(reason='regression in upstream calc, see #123')
def lookup_handles_blank_keys():
    ...
```

`@xfail` accepts the same `reason=` and `condition=` kwargs as
`@skip`. When both decorators are applied, `@skip` wins.

### `@tag`

Attach free-form tags to filter tests at the command line. Multiple
`@tag` decorators stack (set-union); a class-level `@tag` propagates
to every method on the class.

```python
from testsweet import tag, test


@test
@tag('slow')
@tag('libreoffice')
def full_recalc_of_quarterly_model():
    ...
```

Filter at run time with `-t` / `--tag` and `-T` / `--exclude-tag`
(both repeatable):

```bash
uv run python -m testsweet -t slow -T flaky tests/
```

### Fixtures

testsweet has no fixture system of its own. For function-style
tests, use any context manager:

```python
from contextlib import contextmanager

from testsweet import test

from sheetwright.testing import Model


@contextmanager
def fresh_model():
    m = Model.open('.')
    try:
        yield m
    finally:
        pass  # nothing to tear down; placeholder for resources


@test
def revenue_grows():
    with fresh_model() as m:
        m.set('growth_rate', 0.05)
        assert math.isclose(
            m.get('Outputs!B1'), 1_050_000, rel_tol=1e-9,
        )
```

For shared per-class state (the equivalent of unittest's
`setUpClass` / `tearDownClass`), implement the context-manager
protocol on the class. The runner enters it for the duration of the
class's method calls — handy for paying the LibreOffice cost once on
a baseline workbook:

```python
from contextlib import AbstractContextManager

from testsweet import test

from sheetwright.testing import Model


@test
class WarmModel(AbstractContextManager):
    def __enter__(self):
        self.model = Model.open('.')
        self.model.recalc()  # pay the LibreOffice cost once
        return self

    def __exit__(self, exc_type, exc, tb):
        return None

    def baseline_revenue(self):
        assert math.isclose(
            self.model.get('Outputs!B1'), 1_040_000, rel_tol=1e-9,
        )

    def named_range_resolves(self):
        assert self.model.workbook.named_ranges[0].name == 'growth_rate'
```

For per-method setup/teardown (the equivalent of `setUp` /
`tearDown`), define `__test_context__` on the class. The runner
enters it once per test method, inside the class's
`__enter__` / `__exit__` scope:

```python
from contextlib import contextmanager


@test
class RevenueScenarios(AbstractContextManager):
    def __enter__(self):
        self.model = Model.open('.')
        return self

    def __exit__(self, exc_type, exc, tb):
        return None

    @contextmanager
    def __test_context__(self):
        # reset assumptions between methods
        self.model.set('growth_rate', 0.04)
        yield

    def baseline(self):
        assert math.isclose(
            self.model.get('Outputs!B1'), 1_040_000, rel_tol=1e-9,
        )

    def aggressive_growth(self):
        self.model.set('growth_rate', 0.10)
        assert math.isclose(
            self.model.get('Outputs!B1'), 1_100_000, rel_tol=1e-9,
        )
```

Subclasses can chain a parent's `__test_context__` via
`super().__test_context__()`.

## A complete example

`tests/test_revenue.py`:

```python
import math

from testsweet import params, test

from sheetwright.testing import Model


@test
def baseline_revenue_at_default_growth_rate():
    model = Model.open('.')
    assert math.isclose(model.get('Outputs!B1'), 1_040_000, rel_tol=1e-9)


@test
@params([
    (0.00, 1_000_000),
    (0.05, 1_050_000),
    (0.10, 1_100_000),
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
sheetwright test
```

Or invoke testsweet directly:

```bash
uv run python -m testsweet tests/        # run a directory
uv run python -m testsweet               # use configured discovery
uv run python -m testsweet tests/test_revenue.py::revenue_scales_with_growth_rate
```

The runner prints one line per test and exits non-zero if any test
fails. Outcomes are returned from `testsweet.run()` as one of
`Passed`, `Failed`, `Errored`, `Skipped`, `XFailed`, or `XPassed`
— see the [testsweet reference](https://github.com/kaapstorm/testsweet/blob/main/docs/reference.md)
for the full sum type.

## Discovery configuration

Configure discovery in `pyproject.toml`:

```toml
[tool.testsweet.discovery]
include_paths = ['tests']
exclude_paths = ['tests/fixtures']
test_files = ['test_*.py', '*_test.py']
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
