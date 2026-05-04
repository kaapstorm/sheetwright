# Importing an existing workbook

Onboard an existing `.xlsx` into a sheetwright project.

## 1. Prepare a directory

```bash
mkdir my-model
cd my-model
```

`sheetwright import` creates the project files (`sheetwright.toml`,
`workbook.toml`, `sheets/`, `data/`) at the target path; it does
not need `init` to have run first.

## 2. Import

```bash
$ sheetwright import path/to/legacy-model.xlsx --archive
Imported path/to/legacy-model.xlsx into /home/you/my-model
```

Flags:

- `--archive` — copies the xlsx into `imports/<timestamp>_<name>.xlsx`
  so you have an audit trail of what landed when.
- `--flatten` — replaces external-reference formulas (e.g.
  `'[other.xlsx]Sheet1'!A1`) with their cached calculated values.
  Required if your workbook links to other workbooks; see
  [external references](#external-references) below.

After import, the project layout is fully populated:

```
my-model/
├── sheetwright.toml
├── workbook.toml
├── sheets/
│   ├── 01_assumptions.md
│   ├── 01_assumptions.yaml
│   ├── 02_outputs.md
│   └── 02_outputs.yaml
├── data/
└── imports/
    └── 2026-04-30T10-12-00_legacy-model.xlsx
```

See [source format](../reference/source-format.md) for what each
file contains.

## 3. Commit

```bash
git init
git add .
git commit -m "Import legacy-model.xlsx"
```

`build/` and `.sheetwright/` are gitignored by default. `imports/`
is committed so reviewers can see the source xlsx alongside the
derived text.

## 4. Verify the round-trip

Build the xlsx back and diff against the original:

```bash
$ sheetwright build
Built build/my-model.xlsx

$ sheetwright diff --vs xlsx:imports/2026-04-30T10-12-00_legacy-model.xlsx
```

Exit code 0 means the round-trip is lossless: source → xlsx is the
inverse of xlsx → source. Any output describes cells that
serialised differently — usually a sign of a feature sheetwright
doesn't support (yet) or a quirk in the original xlsx (e.g.
zero-width whitespace in a string).

## 5. First test

Pick a single-cell assertion to anchor your understanding of the
model:

```python
# tests/test_smoke.py
import math

from testsweet import test

from sheetwright.testing import Model


@test
def revenue_matches_imported_workbook():
    model = Model.open('.')
    assert math.isclose(
        model.get('Outputs!B5'), 1_234_567, rel_tol=1e-6
    )
```

Run it:

```bash
$ sheetwright test
tests/test_smoke.py::revenue_matches_imported_workbook ... ok
```

If this passes, your import is faithful at least at this cell — a
useful baseline before you start refactoring.

## 6. Snapshot the rest

```bash
$ sheetwright snapshot
initialized snapshot at tests/snapshots/my-model.json
```

The snapshot covers *every* formula cell. From now on, any change
that perturbs a calculated value shows up in
`sheetwright snapshot` output. See [snapshots
tutorial](snapshots.md).

## External references

If your workbook contains formulas that read from other workbooks,
import refuses by default:

```
$ sheetwright import legacy.xlsx
Error: Workbook contains external references; pass --flatten to
replace them with cached values, or resolve them in Excel before
importing.
First few: '[other.xlsx]Sheet1'!A1, '[refs.xlsx]Lookup'!B5:B100, ...
```

Two options:

1. **Resolve in Excel first.** Open the workbook in Excel, paste-
   special the linked ranges as values, save, and re-import. This
   is the cleanest path; the resulting source is self-contained.
2. **`--flatten`.** Tell sheetwright to replace external-reference
   formulas with their last cached values:

   ```bash
   sheetwright import legacy.xlsx --flatten --archive
   ```

   The structural shape of formulas that *internally* used the
   external value is preserved; only the leaf reference is
   replaced. Local formulas are unchanged.

## Re-importing after edits

If a colleague edits `build/my-model.xlsx` directly in Excel and
saves changes, sheetwright won't lose those edits — but it does
make you reconcile them. The re-import flow detects the divergence,
shows a diff, and asks how to merge. See [escape hatch](escape-hatch.md)
for that workflow.

## Where to next?

- [Source format](../reference/source-format.md) — understand what
  the import produced.
- [TDD workflow](tdd-workflow.md) — start adding tests.
- [Escape hatch](escape-hatch.md) — for the next round-trip when
  someone edits the xlsx directly.
