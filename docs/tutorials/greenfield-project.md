# Greenfield project

Build a small revenue model from scratch. This tutorial walks
through the full loop: scaffold, write source, build, recalc, test,
snapshot.

## 1. Scaffold

```bash
$ sheetwright init my-revenue-model
Initialised sheetwright project at /home/you/my-revenue-model
$ cd my-revenue-model
```

The default scaffold gives you `sheetwright.toml`, `workbook.toml`,
empty `sheets/` and `data/` directories, an empty `tests/__init__.py`,
and a `.gitignore`. See [source format](../reference/source-format.md)
for the full layout.

## 2. Plan the workbook

Two sheets:

- `Assumptions` — inputs the analyst tweaks.
- `Outputs` — derived values (revenue, growth, sensitivity).

Edit `workbook.toml`:

```toml
[workbook]
name = "my-revenue-model"
sheets = ["Assumptions", "Outputs"]

[[named_ranges]]
name = "growth_rate"
scope = "workbook"
ref = "Assumptions!$B$1"

[[named_ranges]]
name = "base_revenue"
scope = "workbook"
ref = "Assumptions!$B$2"
```

## 3. Write `sheets/01_assumptions.md`

```markdown
| (cell) | A             | B       |
| ---    | ---           | ---     |
| 1      | growth_rate   | 0.04    |
| 2      | base_revenue  | 1000000 |
| 3      | tax_rate      | 0.21    |
```

The first column header is the literal string `(cell)`; subsequent
columns are Excel column letters. Numbers, strings, `TRUE`/`FALSE`,
and formulas (starting with `=`) are all accepted. See [source
format](../reference/source-format.md#markdown-table-format-sheetsn_slugmd).

## 4. Write `sheets/02_outputs.md`

```markdown
| (cell) | A           | B                            |
| ---    | ---         | ---                          |
| 1      | revenue_y1  | =base_revenue*(1+growth_rate) |
| 2      | revenue_y2  | =B1*(1+growth_rate)           |
| 3      | tax_y1      | =B1*Assumptions!B3           |
| 4      | net_y1      | =B1-B3                       |
```

Note the formulas mix named ranges (`growth_rate`, `base_revenue`)
and explicit sheet references (`Assumptions!B3`). Both work.

## 5. Add formats (optional)

Create `sheets/01_assumptions.yaml` for percentages:

```yaml
formats:
  pct:
    number_format: "0.00%"

cell_formats:
  B1: pct
  B3: pct

column_widths:
  A: 18.0
  B: 12.0
```

The YAML sidecar carries everything that's not values or formulas
— number formats, fonts, fills, validation, comments, conditional
formatting. See [source
format](../reference/source-format.md#yaml-sidecar-sheetsn_slugyaml)
for the full schema.

## 6. Build

```bash
$ sheetwright build
Built /home/you/my-revenue-model/build/my-revenue-model.xlsx
```

Open it in Excel or LibreOffice to confirm. The build is
deterministic — re-running produces a byte-identical file.

## 7. Recalc

```bash
$ sheetwright recalc
recalculated: .sheetwright/calc/4f1a9c2e8b7d.json
```

`recalc` invokes LibreOffice headless against `build/<name>.xlsx`,
collects calculated values, and caches them keyed by the xlsx's
SHA-256. A second `recalc` with no source changes is instant:

```bash
$ sheetwright recalc
cache hit: 4f1a9c2e8b7d
```

See [calc engine](../reference/calc-engine.md) for cache details.

## 8. Lint

```bash
$ sheetwright check
no issues
```

`check` flags dangling sheet references, undefined named ranges,
and manifest/filesystem mismatches. Run it whenever you've moved
things around.

## 9. Write a test

`tests/test_revenue.py`:

```python
import math

from testsweet import params, test

from sheetwright.testing import Model


@test
def revenue_y1_at_default_growth_rate():
    model = Model.open('.')
    assert math.isclose(
        model.get('Outputs!B1'), 1_040_000.0, rel_tol=1e-9
    )


@test
@params([
    (0.00, 1_000_000),
    (0.05, 1_050_000),
    (0.10, 1_100_000),
])
def revenue_y1_scales_with_growth(rate, expected):
    model = Model.open('.')
    model.set('growth_rate', rate)
    assert math.isclose(model.get('Outputs!B1'), expected, rel_tol=1e-9)
```

Run:

```bash
$ sheetwright test
tests/test_revenue.py::revenue_y1_at_default_growth_rate ... ok
tests/test_revenue.py::revenue_y1_scales_with_growth[rate=0.0,expected=1000000] ... ok
tests/test_revenue.py::revenue_y1_scales_with_growth[rate=0.05,expected=1050000] ... ok
tests/test_revenue.py::revenue_y1_scales_with_growth[rate=0.1,expected=1100000] ... ok
```

See [testing reference](../reference/testing.md) for the full API.

## 10. Snapshot

```bash
$ sheetwright snapshot
initialized snapshot at tests/snapshots/my-revenue-model.json
```

The snapshot captures every formula cell's calculated value.
Subsequent runs diff against it:

```bash
$ sheetwright snapshot
no changes
```

If you change a formula and the calculated outputs drift:

```bash
$ sheetwright snapshot
  Outputs!B1: 1040000.0 -> 1052000.0
  Outputs!B2: 1081600.0 -> 1106104.0
```

Exit code 1. Either fix the regression, or accept it:

```bash
$ sheetwright snapshot --update
updated snapshot at tests/snapshots/my-revenue-model.json
```

See [snapshots tutorial](snapshots.md) for the full workflow.

## 11. Commit

```bash
git init
git add sheetwright.toml workbook.toml sheets/ tests/ .gitignore
git commit -m "Initial revenue model"
```

`build/` and `.sheetwright/` are ignored by default.

## 12. Hand off the xlsx

When you're ready to share with non-sheetwright users:

```bash
sheetwright build
cp build/my-revenue-model.xlsx /path/to/share/
```

Or commit `build/<name>.xlsx` selectively (override the gitignore
with `git add -f`) if your team prefers to pull from git. Just
remember the canonical source is `sheets/`; the `recalc` and
`snapshot` commands warn if the committed xlsx diverges from
source.

## Where to next?

- [TDD workflow](tdd-workflow.md) — write the test before the
  formula.
- [Importing existing](importing-existing.md) — onboard an existing
  xlsx instead of starting fresh.
- [Escape hatch](escape-hatch.md) — handle the case where someone
  edits the built xlsx directly in Excel.
