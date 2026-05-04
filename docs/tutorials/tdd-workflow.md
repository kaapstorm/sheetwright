# Test-driven model development

Write the failing test first; then make the model satisfy it.
Spreadsheets benefit from this even more than code: stakeholders
look at the rendered output and rarely catch a sign error in a
nested IF, but a one-line test can.

## Why TDD for spreadsheets

Three things go wrong with models in production:

1. **Stale formulas.** A row insert breaks `=SUM(B2:B100)`. The
   error is silent.
2. **Sign / unit errors.** `=B1*0.21` vs `=B1*(1-0.21)` — both
   render as a number; only one is right.
3. **Refactor regressions.** A name change ripples; one place is
   missed.

Tests catch all three. The contract is the same as for code: a test
declares *what should be true*; refactors must preserve it.

## The loop

1. Write the test against an output cell. Don't write the formula
   yet.
2. Run `claudesheets test`. Watch it fail.
3. Edit the source (`sheets/<n>_<slug>.md` or `.yaml`).
4. Run `claudesheets build && claudesheets test`. Watch it pass.
5. Commit.

## Worked example

Imagine a model that needs a new column: post-tax profit. We don't
have it yet.

### Step 1: write the failing test

`tests/test_post_tax_profit.py`:

```python
import math

from testsweet import test, test_params

from claudesheets.testing import Model


@test
def post_tax_profit_at_default_inputs():
    model = Model.open('.')
    # revenue 1_000_000 * (1 + 0.04) = 1_040_000; tax 21%
    expected = 1_040_000 * (1 - 0.21)
    assert math.isclose(
        model.get('Outputs!C1'), expected, rel_tol=1e-9
    )


@test_params([
    {'rate': 0.00, 'tax': 0.21, 'revenue': 1_000_000},
    {'rate': 0.05, 'tax': 0.21, 'revenue': 1_050_000},
    {'rate': 0.05, 'tax': 0.30, 'revenue': 1_050_000},
])
def post_tax_profit_scales(rate, tax, revenue):
    model = Model.open('.')
    model.set('growth_rate', rate)
    model.set('Assumptions!B3', tax)
    expected = revenue * (1 - tax)
    assert math.isclose(
        model.get('Outputs!C1'), expected, rel_tol=1e-9
    )
```

### Step 2: run, watch it fail

```bash
$ claudesheets test
tests/test_post_tax_profit.py::post_tax_profit_at_default_inputs ... ERROR: RuntimeError: calc engine omitted formula cell Outputs!C1
```

`Model.get` raises when a formula cell is in the source but missing
from the calc result. In this case there's no formula at all yet,
and the literal cell is empty. The test fails for the right reason.

### Step 3: edit the source

Add column C to `sheets/02_outputs.md`:

```markdown
| (cell) | A           | B                            | C                       |
| ---    | ---         | ---                          | ---                     |
| 1      | revenue_y1  | =base_revenue*(1+growth_rate) | =B1*(1-Assumptions!B3) |
```

### Step 4: rebuild and retest

```bash
$ claudesheets build && claudesheets test
Built build/my-revenue-model.xlsx
tests/test_post_tax_profit.py::post_tax_profit_at_default_inputs ... ok
tests/test_post_tax_profit.py::post_tax_profit_scales[rate=0.0,tax=0.21,revenue=1000000] ... ok
tests/test_post_tax_profit.py::post_tax_profit_scales[rate=0.05,tax=0.21,revenue=1050000] ... ok
tests/test_post_tax_profit.py::post_tax_profit_scales[rate=0.05,tax=0.3,revenue=1050000] ... ok
```

### Step 5: commit

```bash
git add sheets/02_outputs.md tests/test_post_tax_profit.py
git commit -m "Add post-tax profit column with tests"
```

The diff in code review is two changes: the new column in
`sheets/02_outputs.md` (the *what*) and the new test (the *why*).
A reviewer can verify both at the formula level.

## Working with Claude Code

Claude Code runs the loop happily. A productive prompt:

> Add a "post-tax profit" column to `sheets/02_outputs.md` as
> column C. Write a test in `tests/test_post_tax_profit.py` that
> asserts the value at `Outputs!C1` when growth rate is 0% and tax
> rate is 21%. Run `claudesheets test` and iterate until it passes.

Claude can:

- Read your `sheets/` and `workbook.toml` to understand the layout.
- Write the test in the right shape (testsweet `@test` /
  `@test_params`, `Model.open('.')`).
- Edit the markdown table and rebuild.
- Run `claudesheets test` and react to the output.

The text-source format makes all of this normal-code work for
Claude — it's reading and editing files, not tickling an Excel API.

## Tips

- **Anchor on outputs, vary inputs.** Tests read best when they say
  "for this input, the output is X." `model.set` makes this cheap.
- **Use named ranges for inputs you'll vary.** `model.set('growth_rate', 0.05)`
  is more durable than `model.set('Assumptions!B1', 0.05)` when the
  layout shifts.
- **Use `math.isclose` for floats.** Calc engines round in the last
  bit. `rel_tol=1e-9` is fine for most accounting math; tighten for
  numeric work.
- **Snapshot the rest.** Tests are for the assertions you care
  about; snapshots catch the ones you didn't think of. Run both.

## See also

- [Testing reference](../reference/testing.md) — full `Model` API.
- [Snapshots tutorial](snapshots.md) — golden-file regression for
  every formula cell.
- [Greenfield project](greenfield-project.md) — the underlying
  build / recalc / test loop.
