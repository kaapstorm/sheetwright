# Snapshots

Golden-file regression for calculated outputs. A snapshot captures
every formula cell's value; a future change that perturbs any of
them surfaces in `sheetwright snapshot` output.

## When to use snapshots

- After a refactor: "no calculated value should change."
- In CI: assert that the committed source produces the committed
  outputs, byte-for-byte.
- During exploratory work: see at a glance which cells your latest
  edit moved.

## What's covered

Snapshots cover **formula cells only**. Literal inputs are
deliberately excluded — the snapshot is about *derived* values, so
edits to assumptions don't drown out actual model changes. If you
want to assert on inputs too, use a regular test.

The snapshot lives at `tests/snapshots/<workbook-name>.json`. It's
checked into git.

## Initial snapshot

The first run creates the file:

```bash
$ sheetwright snapshot
initialized snapshot at tests/snapshots/my-revenue-model.json
```

The JSON is sorted by sheet then address, with one entry per
formula cell:

```json
{
  "Outputs": {
    "B1": 1040000.0,
    "B2": 1081600.0,
    "C1": 821600.0
  }
}
```

Commit it:

```bash
git add tests/snapshots/my-revenue-model.json
git commit -m "Initial snapshot"
```

## Subsequent runs

```bash
$ sheetwright snapshot
no changes
```

Exit code 0 — the calculation matches the committed snapshot.

When something changes:

```bash
$ sheetwright snapshot
  Outputs!B1: 1040000.0 -> 1052000.0
  Outputs!B2: 1081600.0 -> 1106104.0
  Outputs!C1: 821600.0 -> 831080.0
```

Exit code 1. Three things might be happening:

1. **You changed an assumption.** The growth rate moved from 4% to
   5.2%. The model is correct; the snapshot is stale. Update.
2. **You changed a formula deliberately.** A new tax treatment.
   Same story: snapshot is stale, update.
3. **You broke something.** The revenue formula now references the
   wrong cell. Don't update — fix the source.

The diff alone doesn't tell you which case you're in. Read it,
decide, and either revert your change or accept the new outputs.

## Reviewing the diff

The CLI prints a flat list. For larger diffs, review the JSON itself:

```bash
$ sheetwright snapshot       # see what's different
$ sheetwright snapshot --update
$ git diff tests/snapshots/   # full structured view in the diff
```

Then *decide before committing*. The snapshot file is just JSON;
git diff treats it as text.

## Accepting a change

```bash
$ sheetwright snapshot --update
updated snapshot at tests/snapshots/my-revenue-model.json
```

This overwrites the saved file with the current calculation.
Always inspect with `git diff` before committing.

## In CI

```yaml
# .github/workflows/check.yml
- run: sheetwright build
- run: sheetwright recalc
- run: sheetwright snapshot
- run: sheetwright test
```

`snapshot` exits 1 on diffs, which fails the job. Combined with
`test`, you cover both targeted assertions and the everything-else
of formula evaluation.

## How it works

Under the hood:

1. `snapshot` builds (or reuses cached) calculated values via the
   calc engine.
2. It walks the source workbook and records which cells are
   formulas.
3. It filters the calc result to keep only formula-cell values.
4. Datetime values are normalised to ISO-8601 strings (so JSON
   round-trips preserve equality).
5. It compares to the saved file on disk.

The cache is shared with `sheetwright recalc`: if you've already
recalc'd, snapshot is instant.

## Pairing with tests

Snapshots and tests do different jobs:

- **Tests** assert specific behaviour: "with growth rate 5%, year-1
  revenue is X." Readable, intentional, focused.
- **Snapshots** catch *unintended* changes anywhere. Comprehensive,
  unselective.

A healthy project uses both. Tests document what matters; snapshots
make sure nothing else moved.

## See also

- [CLI reference: snapshot](../reference/cli.md#snapshot) — flags
  and exit codes.
- [Testing reference](../reference/testing.md) — for targeted
  assertions.
- [Calc engine](../reference/calc-engine.md) — the underlying
  recalc.
