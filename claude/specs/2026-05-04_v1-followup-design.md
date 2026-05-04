# sheetwright v1 — review follow-up design

**Status:** draft
**Reference review:** [`claude/reviews/2026-05-04_v1-codebase-review.md`](../reviews/2026-05-04_v1-codebase-review.md)

## Purpose

Sequence the 35 findings from the v1 review into coherent work
phases. Bundle by root cause rather than by severity alone — three
of the major findings dissolve into one structural change, and most
minors fall out as side effects of those structural changes.

## Phasing principles

- **Security fixes ship first** and are not bundled with refactors.
- **Structural fixes precede cosmetic ones** so we don't re-touch the
  same files twice.
- **Each phase is independently mergeable** with a green test suite.
- A phase is a *plan* (lives under `claude/plans/`) once it's
  detailed; this document is the index.

## Phases

### Phase 1 — Security hardening (urgent)

Goal: close the critical and major security findings before any
non-trusted MCP client is plausible.

| Findings   | Change                                                                                                                                                                                                                                                                                             |
|------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **F1**, F6 | Treat each tool call's `project` argument as the workspace root for that call. Every other path on the same call (`xlsx`, `out_path`, `vs`, `path`, …) must resolve under the project root after `Path.resolve()`; reject otherwise with `path_outside_project`. `do_test` stops importing arbitrary paths: test discovery is constrained to `<project>/tests/`. Preserves the documented "one server, many projects" model. |
| F7         | Add a `safe_load_workbook()` helper in `xlsx/` that wraps `openpyxl.load_workbook` with: max uncompressed size, max sheet count, max row × col, max shared-string count. All three current call-sites (reader, flatten, libreoffice) route through it. Limits configurable via `sheetwright.toml`. |
| F8         | Validate / quote SQL identifiers in `bulk._create_table`. Reject any column header that's not `[A-Za-z_][A-Za-z0-9_]*`; double-quote and escape internal `"` in the emitted DDL.                                                                                                                   |
| F19        | Add `--safe-mode --norestore --nolockcheck --nofirststartwizard --nodefault` to the LibreOffice argv in `calc/libreoffice.py`.                                                                                                                                                                     |
| F20        | `apply_session` re-resolves the staged xlsx path against `<workspace>/.sheetwright/staged/` rather than trusting the on-disk `xlsx_path` field.                                                                                                                                                    |
| F35        | Add a `# Trust model` block to `mcp/server.py` and a section to `docs/reference/mcp.md`: trusted operator, untrusted xlsx inputs, and the per-call project-root invariant. Spell out the security implication of the "one server, many projects" choice — a tool call's `project` arg defines the sandbox for that call only; the operator (not the server) is responsible for which project paths the client may choose, since any project on disk is reachable. Update the wiring example to reflect this. |

Deliverables: one plan, one PR. Tests: zip-bomb and traversal
fixtures live under `tests/security/`.

### Phase 2 — Typed command results (the structural fix)

Goal: dissolve findings F2, F4, F16, and unblock F30 / F34. Replicates
the `stage_reimport` / `commit_staged` pattern across every command.

**Shape:**

```python
# commands/build_cmd.py
@dataclass(frozen=True)
class BuildResult:
    xlsx_path: Path
    sheets_written: int

def run(project: Project, *, out: Path | None = None) -> BuildResult: ...
```

**Steps:**

1. Define `SheetwrightError` hierarchy in `exceptions.py` with stable
   string `code` attributes (`project_not_found`, `no_built_xlsx`,
   `dirty_workspace`, `external_ref`, `recalc_failed`, …). Replace
   the closed `MCP error code` set in `mcp/errors.py`.
2. Rewrite each `commands/*.run()` to return a typed `Result` and
   raise `SheetwrightError` subclasses. Stop raising
   `ClickException` from `reimport/flow.py` and `diff/loaders.py`.
3. The Click wrappers (`cli.py` group) become thin: call `run()`,
   `click.echo` a human-readable rendering of the result, translate
   `SheetwrightError → ClickException` at the boundary.
4. The MCP wrappers (`mcp/server.py`) call `run()` directly and
   serialise the `Result` to JSON. Delete `_run_capturing`. Delete
   `classify_click_error`. The exit-code-as-`has_diffs` hack in
   `do_diff` goes away — `DiffResult.has_diffs` is a real field.
5. Hoist the repeated `Project.open` + error-translation block
   (F16) into a single helper used by both the CLI and MCP wrappers.
6. Add a `tests/test_cli_mcp_parity.py` (F34) that asserts every
   command has both surfaces and that the JSON shape of each MCP
   tool matches a golden schema.

Side effects: F30 (structured logging) becomes trivial — introduce
a `sheetwright.log` channel that the MCP wrapper routes to stderr
while the CLI keeps `click.echo` for stdout. F22 (lazy imports)
mostly disappears as cycles dissolve.

### Phase 3 — Conditional-format polymorphism

Goal: kill findings F3 and most of F14.

Each `ConditionalFormat` subclass gains four classmethods/methods —
`to_yaml`, `from_yaml`, `to_xlsx`, `from_xlsx` — registered against a
single `KIND` string. The five-way isinstance ladders in
`source/yaml_sidecar.py` and `xlsx/cf_translate.py` are replaced by
a `_REGISTRY[kind].from_yaml(...)` lookup. Centralise default
constants (`'FF638EC6'`, `'3TrafficLights1'`, etc.) on the dataclass.

Side effects: removes ~10 of the ~20 `# type: ignore` markers.

### Phase 4 — openpyxl seam containment

Goal: kill the rest of F14, plus F13 and F23.

1. Add `xlsx/_compat.py` with thin typed wrappers around the bits of
   openpyxl we use (`column_index_from_string`, `get_column_letter`,
   `Tokenizer`, the `Color` constructor). `source/markdown.py` and
   `diff/check.py` import from there, not from `openpyxl`.
2. Either ship a `xlsx/openpyxl.pyi` stub or replace the remaining
   ignores with an `Any`-typed local alias and a one-line comment
   explaining which openpyxl class is at the boundary.
3. Name the magic constants: `XLSX_RGB_ALPHA = 'FF'`,
   `CACHE_HASH_CHUNK = 65_536`. (F23)

### Phase 5 — Calc engine: real plug-in interface

Goal: deliver on the README's "swappable" claim. Findings F5 and F21.

1. Replace `evaluate(xlsx_path) -> CalcResult` with an interface that
   doesn't bake in process boundaries — e.g. `evaluate(workbook:
   Workbook) -> CalcResult` plus a default `LibreOfficeEngine` that
   serialises to a temp xlsx internally. In-process / remote engines
   become possible without reshaping the interface.
2. Replace the if/elif factory in `calc/__init__.py` with a registry
   that the LibreOffice engine registers into via entry-point or an
   explicit registration call from `__init__`. Configurable via
   `sheetwright.toml` `[calc] engine = "libreoffice"`.
3. Add a `FakeCalcEngine` for tests so the suite no longer needs a
   real LibreOffice install for the bulk of the cases.
4. Rename the temp-dir prefix `'cshs-calc-'` → `'sheetwright-calc-'`
   (F21). Trivial.

### Phase 6 — Quality refactors

Bundled because each one is small and they all touch the same set of
files Phase 2 just rewrote.

| Finding  | Change                                                                                                                                                            |
|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| F11      | Split `xlsx/reader.read_xlsx` into `_read_sheet` + `_make_cell`.                                                                                                  |
| F12      | Extract `_format_outcome` and a `_user_test_namespace` context manager out of `commands/test_cmd.run`.                                                            |
| F15      | Delete `WorkbookManifest.__eq__`; rely on dataclass default. Add a regression test for the latent bug it was masking.                                             |
| F17, F29 | Add `Project.built_xlsx_path` (cached) and a `cached_config` property. Five command modules drop the duplicated `built = project.build_dir / f'{cfg.name}.xlsx'`. |
| F22      | Hoist remaining lazy imports to module level after Phase 2 collapses cycles.                                                                                      |
| F24      | One `sheet_stem(path) -> str` helper consumed by `source/reader`, `source/writer`, `diff/check`.                                                                  |
| F25      | Delete `ImportError_` and `BuildError`.                                                                                                                           |
| F26      | Promote enum-like string sets to `enum.StrEnum`: `NamedRange.scope`, `CheckIssue.kind`, CF `kind`, MCP error codes.                                               |
| F27      | `snapshot_cmd` calls `recalc_cmd.run()` instead of re-implementing cache-or-evaluate.                                                                             |
| F28      | Codebase-wide sweep: `Optional[X]` → `X                                                                                                                           | None`. Adds a ruff rule to keep it that way. |
| F31      | `_format_id` hashes `json.dumps(asdict(fmt), sort_keys=True)` rather than `repr`.                                                                                 |
| F32      | Replace the `update: bool` parameter on snapshot with two methods (`snapshot.run` / `snapshot.update`), or two subcommands at the CLI layer.                      |
| F33      | Drop the `# for mypy` assert; narrow the type at its source.                                                                                                      |

### Phase 7 — State integrity & tooling discipline

| Finding | Change                                                                                                                                                                                                                                                  |
|---------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| F10     | `reimport/session.load_session`, `build_hash.read_build_hash`, `diff/check._bulk_data_issues` all distinguish "absent" from "corrupt": absent returns `None`, corrupt raises `CorruptStateError` (a `SheetwrightError`). The check command surfaces it. |
| F18     | `gitutil.has_uncommitted_changes`: add a 5s timeout, `logger.warning` on subprocess failure, `not_a_repo` short-circuit, and switch the `.is_dir()` check to one that recognises git worktrees (`git rev-parse --git-dir`).                             |
| F9      | Migrate `tests/` from raw `@test` + ad-hoc `@contextmanager` to pytest-unmagic `@fixture` / `@use` per CLAUDE.md. Build a small `tests/fixtures.py` module with `project`, `built_workbook`, `mcp_session`, etc.                                        |

F9 is the largest task by line count; suggest splitting into
sub-plans by test-file group.

## Out of scope

- The README's v2 list (watch-mode, semantic-diff renderer,
  multi-workbook, alternative calc engines, `imports/` retention).
  This document only covers v1 review fallout.

## Suggested execution order

1. Phase 1 (security) — week 1.
2. Phase 2 (typed results) — weeks 2–3.
3. Phase 3 (CF polymorphism) — week 3.
4. Phases 4 + 6 in parallel — week 4.
5. Phase 5 (calc engine) — week 5.
6. Phase 7 (state + tests) — week 6, with F9 spilling into 7.

Each phase gets its own plan under `claude/plans/`, named
`YYYY-MM-DD_v1-followup-phase-N-<slug>.md`, written immediately
before that phase starts.

## Open questions

- **Workspace root sourcing.** Phase 1 needs a workspace-root
  contract for the MCP server. Options: (a) required `--workspace`
  flag on `sheetwright mcp`; (b) auto-discover from cwd at startup;
  (c) per-tool argument.

  **Resolved: (c).** The documented "one server, many projects"
  model (mcp.md:32) means each tool call carries its own `project`
  arg, and that arg is the sandbox for that call. (a) and (b) would
  silently break that contract. The `--directory` in the Claude
  Desktop wiring example is `uv`'s flag, not a sheetwright concept.
  Phase 1 enforces: every other path on a call must resolve under
  the call's `project` root. The trust boundary stays at the
  operator — they decide which project paths the client may choose.
  This trade-off is documented in `docs/reference/mcp.md` (F35).

- **Test-runner safety.** Once `do_test` is constrained to
  `<project>/tests/`, we still `exec_module` the user's code.

  **Resolved: acceptable under the trusted-operator model.** The
  operator chose the project path; tests in that project run
  in-process. Phase 1's `# Trust model` doc block calls this out
  explicitly. Revisit if the trust model widens (e.g. multi-tenant
  hosting), at which point a subprocess boundary is the path.

- **Enum vs StrEnum vs Literal.** F26 has three viable shapes. Pick
  one project-wide rather than mixing.

  **Resolved: `enum.StrEnum`** project-wide. Phase 6 standardises
  `NamedRange.scope`, `CheckIssue.kind`, CF `kind`, and the new
  `SheetwrightError.code` set on it.
