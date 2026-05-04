# CLI reference

Every claudesheets command, every flag.

The CLI is invoked as `claudesheets` (or `uv run claudesheets`). Most
commands take `--project <path>` to point at a project directory;
the default is `.`. All commands return exit code 0 on success;
`diff` and `check` use a non-zero exit to signal "differences" or
"issues found" — these are not errors.

## `init`

Scaffold an empty claudesheets project.

```
claudesheets init [PATH]
```

| Flag | Description |
| --- | --- |
| `PATH` | Directory to create. Defaults to `.`. Must be empty if it exists. |

Creates `claudesheets.toml`, `workbook.toml`, `sheets/`, `data/`,
`tests/__init__.py`, and a `.gitignore` covering `build/` and
`.claudesheets/`.

```bash
claudesheets init my-model
claudesheets init .
```

## `import`

Read an `.xlsx` into source form, or merge updates into an existing
project.

```
claudesheets import [XLSX] [OPTIONS]
```

| Flag | Description |
| --- | --- |
| `XLSX` | Path to an `.xlsx` file. Required for first import or staging a re-import; omitted with `--apply` / `--abort`. |
| `--archive` | Copy the xlsx into the project's `imports/` directory after a successful import. |
| `--flatten` | Replace external-reference formulas with their cached values. Without this, importing a workbook that contains `[other.xlsx]Sheet1!A1`-style references is rejected. |
| `-I`, `--non-interactive` | Stage the diff and exit. Follow up with `--apply` or `--abort`. |
| `--apply` | Apply a previously-staged re-import. |
| `--abort` | Discard a previously-staged re-import. |
| `--force` | Skip the uncommitted-source guard during re-import. Does **not** bypass the external-reference check or auto-overwrite. |
| `--project PATH` | Project directory. Defaults to `.`. |

The behaviour depends on whether `sheets/` is already populated:

- **First import** (`sheets/` is empty): writes `claudesheets.toml`,
  `workbook.toml`, the `.md`/`.yaml` sidecars, and exits.
- **Re-import** (`sheets/` is non-empty): computes a diff, prints it,
  then either prompts interactively (Merge / Overwrite / Reject) or
  stages the diff under `-I`.

Examples:

```bash
# Initial import
claudesheets import path/to/model.xlsx --archive

# Initial import, dropping cross-workbook links
claudesheets import path/to/model.xlsx --flatten

# Interactive re-import after a colleague edited build/my-model.xlsx
claudesheets import build/my-model.xlsx

# Non-interactive re-import for CI: stage, review, then apply
claudesheets import build/my-model.xlsx -I
git diff   # review the printed diff
claudesheets import --apply
# or
claudesheets import --abort
```

The re-import flow is documented in [escape
hatch](../tutorials/escape-hatch.md).

## `build`

Compile source files into an `.xlsx`.

```
claudesheets build [OPTIONS]
```

| Flag | Description |
| --- | --- |
| `--out PATH` | Output path. Defaults to `build/<workbook-name>.xlsx`. |
| `--project PATH` | Project directory. Defaults to `.`. |

```bash
claudesheets build
claudesheets build --out /tmp/preview.xlsx
```

`build` also runs the bulk-data step: every `data/*.csv` is loaded
into `.claudesheets/bulk.sqlite`. The xlsx writer is deterministic —
re-building from unchanged source produces a byte-identical file.

After a successful build, claudesheets records the xlsx's SHA-256 in
`.claudesheets/build-hash.json`. Other commands use this to detect
external edits to the built file (see
[escape hatch](../tutorials/escape-hatch.md)).

## `recalc`

Run the calc engine against the built xlsx and cache the calculated
values.

```
claudesheets recalc [OPTIONS]
```

| Flag | Description |
| --- | --- |
| `--force` | Ignore the cache and re-run the calc engine. |
| `--project PATH` | Project directory. Defaults to `.`. |

```bash
claudesheets recalc
claudesheets recalc --force
```

The cache is keyed by the SHA-256 of the built xlsx. A change to
source that produces a byte-identical xlsx (for instance, a comment
edit in the YAML sidecar that doesn't affect anything) will hit the
cache. See [calc engine](calc-engine.md) for the cache layout.

If the built xlsx has been edited externally, `recalc` warns to
stderr before running. The warning does not change the exit code.

## `test`

Run the project's testsweet tests.

```
claudesheets test [TARGETS...] [OPTIONS]
```

| Flag | Description |
| --- | --- |
| `TARGETS...` | Specific test files or directories under the project. Defaults to `tests/test_*.py`. |
| `--project PATH` | Project directory. Defaults to `.`. |

```bash
claudesheets test
claudesheets test tests/test_revenue.py
claudesheets test tests/integration/
```

`test` discovers `tests/test_*.py` files and imports each under a
stable `_user_tests.<stem>` namespace. It does **not** read
`[tool.testsweet.discovery]` from your `pyproject.toml`; if you want
that, run `python -m testsweet` directly from the project root.

Exit code 1 if any test failed, errored, or `XPASS`-ed.

See [testing reference](testing.md) for the test-author API.

## `snapshot`

Compare or update the golden-file snapshot of calculated values.

```
claudesheets snapshot [OPTIONS]
```

| Flag | Description |
| --- | --- |
| `--update` | Overwrite the saved snapshot with current values. |
| `--project PATH` | Project directory. Defaults to `.`. |

```bash
claudesheets snapshot           # diff vs saved snapshot
claudesheets snapshot --update  # overwrite the saved snapshot
```

Snapshots cover formula cells only — literal inputs are intentionally
excluded so a snapshot diff highlights actual model changes, not
input edits. The first run initialises the snapshot under
`tests/snapshots/<workbook-name>.json`. Subsequent runs print a diff
and exit 1 if anything differs.

See [snapshots tutorial](../tutorials/snapshots.md).

## `diff`

Show a semantic diff between source and a target workbook.

```
claudesheets diff [OPTIONS]
```

| Flag | Description |
| --- | --- |
| `--vs TARGET` | Comparison target: `xlsx:<path>` or `source:<path>`. Default: `build/<name>.xlsx` of this project. |
| `--project PATH` | Project directory. Defaults to `.`. |

```bash
claudesheets diff
claudesheets diff --vs xlsx:other.xlsx
claudesheets diff --vs source:../other-project
```

Exit code 0 if there are no differences, 1 otherwise. Useful in CI
to assert "the built xlsx matches source" or "two projects agree."

## `check`

Lint dangling refs, missing names, and schema mismatches.

```
claudesheets check [OPTIONS]
```

| Flag | Description |
| --- | --- |
| `--project PATH` | Project directory. Defaults to `.`. |

```bash
claudesheets check
```

Exit code 0 if there are no issues, 1 otherwise. Issue kinds:

- `dangling_sheet_ref` — formula references an unknown sheet.
- `dangling_named_range` — formula uses an undefined named range.
- `manifest_sheet_missing` — `workbook.toml` lists a sheet without a
  `sheets/` file.
- `sheet_file_missing_from_manifest` — a `sheets/*.md` file isn't in
  the manifest.
- `bulk_data_missing` — a sheet's `source:` directive points at a
  missing CSV.

## `mcp`

Run the MCP server on stdio.

```
claudesheets mcp
```

No flags. The process stays alive until the client closes the
connection. Most MCP clients launch this as a subprocess. See the
[MCP reference](mcp.md) for tool documentation and a wiring example.

## Exit codes summary

| Command | 0 | 1 |
| --- | --- | --- |
| `init`, `import`, `build`, `recalc`, `test` | success | error |
| `test` | all tests passed | any failure / error / XPASS |
| `diff` | no differences | differences found |
| `check` | no issues | issues found |
| `snapshot` | matches saved snapshot (or initialised) | diffs found |
| `mcp` | client disconnected cleanly | error |

`check` and `diff` use exit 1 as a *signal*, not an error. CI scripts
should branch on the exit code to decide whether to fail the build.
