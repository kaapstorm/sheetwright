# Project conventions

## Layout

- `src/claudesheets/` — library + CLI entry points
- `tests/` — pytest tests, mirroring the package layout
- `claude/specs/` — design specifications
- `claude/plans/` — implementation plans (when written)

## Design specs and plans

Design specifications live in `claude/specs/`. Implementation plans live
in `claude/plans/`. Both use the filename convention
`YYYY-MM-DD_name-of-document.md`. Use the date the document was first
written.

## Python and tooling

- Python 3.11 (`.python-version`)
- [`uv`](https://docs.astral.sh/uv/) is the project manager — use
  `uv sync` to install, `uv run <cmd>` to run, `uv add <pkg>` to add
  runtime deps, `uv add --dev <pkg>` for dev deps. Don't edit
  `pyproject.toml` dependency lists by hand unless you also know to
  update `uv.lock`.
- [`ruff`](https://docs.astral.sh/ruff/) is the formatter/linter.
  **Run `uv run ruff format <changed-files>` before every commit.**
  Project style is `line-length = 79` and `quote-style = 'single'` —
  use single quotes in new code; ruff format will fix mixed quoting.
- [`mypy`](https://mypy.readthedocs.io/) checks type annotations.
  **Run `uv run mypy src/` before committing changes that add or
  modify type annotations.** Public functions and CLI entry points
  must have type hints; mypy must pass before commit.

## Testing

Use [`pytest`](https://docs.pytest.org/) with
[`pytest-unmagic`](https://github.com/kaapstorm/pytest-unmagic/tree/nh/docs_5/)
(branch `nh/docs_5` for current docs). pytest-unmagic replaces pytest's
implicit fixture system with explicit imports — fixtures are imported and
called like normal functions, which makes dependencies visible and helps
both humans and Claude reason about what a test depends on.

### Conventions

**Define fixtures with `@fixture`, yield exactly once.** Setup before
yield, teardown after.

```python
from unmagic import fixture

@fixture
def workbook():
    wb = build_test_workbook()
    yield wb
    wb.close()
```

**Apply fixtures explicitly with `@use(...)` or the shorthand
`@<fixture>`.** Do not rely on parameter-name matching.

```python
from unmagic import use
from claudesheets.testing.fixtures import workbook

@use(workbook)
def test_recalc_updates_outputs():
    wb = workbook()
    wb.set("Assumptions!growth_rate", 0.05)
    wb.recalc()
    assert wb.get("Outputs!revenue_2027") == 1_234_567
```

The shorthand `@workbook` is equivalent to `@use(workbook)` for a single
fixture — fine to use when there's no ambiguity.

**Set scope explicitly when fixtures are expensive.** Building or
recalculating a workbook is slow; share where the test contract allows
it.

```python
@fixture(scope="module")
def calc_engine():
    engine = LibreOfficeEngine.start()
    yield engine
    engine.stop()
```

**Autouse via `@fixture(autouse=__file__)`** when a test file needs
shared setup/teardown for every test in it.

**`unittest.TestCase` is supported** — pytest-unmagic fixtures work on
class-based tests, unlike standard pytest fixtures. Prefer plain
functions, but reach for `TestCase` when behaviour is genuinely
class-shaped.

### Don't

- Don't use bare `pytest.fixture` — always use `unmagic.fixture` so the
  explicit-import contract is uniform across the codebase.
- Don't put shared fixtures in `conftest.py` and rely on pytest's
  auto-discovery to wire them up. Put them in a regular module
  (e.g. `claudesheets.testing.fixtures` or `tests/fixtures.py`) and
  import them where used.
- Don't pass fixture names as parameters expecting pytest to inject
  them — that's the magic pytest-unmagic exists to avoid.

## Code style

- Match the surrounding code. The project is young; when in doubt,
  prefer the simplest readable option.
- Type hints on public functions and CLI entry points.
- No ad-hoc comments explaining what code does. Comment only when a
  non-obvious *why* matters.
