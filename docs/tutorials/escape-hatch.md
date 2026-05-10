# Escape hatch: handling external xlsx edits

Sometimes someone edits `build/<name>.xlsx` directly. A colleague
opens it in Excel, drags a few values around, hits save. Or you do.
That edit is now in the build artefact but not in `sheets/`, and a
`sheetwright build` would silently overwrite it.

sheetwright detects this and gives you a structured way to merge
the edit back into source.

## 1. Detection

The first command after the external edit warns:

```
$ sheetwright recalc
WARNING: build/my-model.xlsx has been modified externally.
Run `sheetwright import build/my-model.xlsx` to review changes.
recalculated: .sheetwright/calc/...
```

The warning goes to stderr; the command itself still runs (so you
can poke around before deciding what to do).

Detection works because `sheetwright build` records the SHA-256 of
the produced xlsx in `.sheetwright/build-hash.json`. Any difference
on a later command is the trigger.

## 2. Re-import (interactive)

```bash
$ sheetwright import build/my-model.xlsx
```

Output: a printed diff of every cell that changed (values,
formulas, formats), then a prompt:

```
Apply changes? [m]erge / [o]verwrite / [r]eject:
```

| Choice                | Effect                                                                                               |
|-----------------------|------------------------------------------------------------------------------------------------------|
| `m` (merge)           | Apply the diff to source. Existing formulas/formats unchanged unless the xlsx changed them.          |
| `o` (overwrite)       | Same as merge in current implementation; the `m`/`o` distinction is reserved for future granularity. |
| `r` (reject, default) | Source unchanged. The xlsx still differs from what source would produce.                             |

If you reject, the warning will keep firing on subsequent commands
until you either re-run import and merge, or rebuild from source
(which restores the build hash):

```bash
sheetwright build   # overwrites build/<name>.xlsx with the source view
```

## 3. Re-import (non-interactive, for CI / Claude)

```bash
$ sheetwright import build/my-model.xlsx -I
... full diff printed ...
Run `sheetwright import --apply` to apply, or `--abort` to discard.
```

`-I` (`--non-interactive`) stages the diff and exits. The session
file is `.sheetwright/reimport.json`. Follow up:

```bash
sheetwright import --apply   # accept the staged diff
# or
sheetwright import --abort   # discard
```

Between staging and applying, the staged xlsx is hashed; if it's
been modified again you get an error and have to re-stage:

```
Error: Staged xlsx at build/my-model.xlsx has been modified since
`-I` (recorded hash 4f1a9c2e8b7d, current 9c2a8f4b1e7d). Re-stage
with `-I` to see the new diff.
```

This is the same flow Claude Code takes via the MCP server, where
the three steps are `do_reimport_stage`, `do_reimport_apply`, and
`do_reimport_abort`. See [MCP reference](../reference/mcp.md).

## 4. Uncommitted-source guard

If your `sheets/` directory has uncommitted git changes,
re-importing refuses:

```
Error: sheets/ has uncommitted changes. Commit or stash before
re-importing, or pass --force.
```

The reasoning: a re-import overwrites source files. Stacking that
on top of in-progress edits makes it hard to recover if you don't
like the merge. Commit (or stash) your work first, then re-import,
then resolve any merge conflict in normal git.

To bypass:

```bash
sheetwright import build/my-model.xlsx --force
```

`--force` skips *only* the uncommitted-source guard. It does **not**
bypass the external-references check (use `--flatten` for that)
and does **not** auto-overwrite without prompting — you still get
the merge / overwrite / reject prompt (or the `-I` staging flow).

## 5. Archiving the xlsx

```bash
sheetwright import build/my-model.xlsx --archive
```

`--archive` copies the xlsx into `imports/<timestamp>_<name>.xlsx`
after a successful import. Useful when the external edit came from
a colleague — you keep an audit trail of who-edited-what-when. The
`imports/` directory is checked into git by default.

## 6. The full sequence

A realistic workflow:

```bash
# Colleague has edited build/my-model.xlsx and emailed it back.
cp ~/Downloads/my-model.xlsx build/my-model.xlsx

# 1. Verify what they changed.
sheetwright diff --vs xlsx:build/my-model.xlsx

# 2. Stage the diff for review.
sheetwright import build/my-model.xlsx -I --archive

# 3. Inspect the staged diff (rendered above; also
#    saved in .sheetwright/reimport.json).

# 4a. Accept.
sheetwright import --apply
git add sheets/ imports/ workbook.toml
git commit -m "Merge edits from colleague"

# 4b. Or reject.
sheetwright import --abort
```

## 7. When you don't want this

If your team is happy with sheets/ being authoritative and the
xlsx being throwaway: never edit `build/`. The warning is the
seatbelt; if you ignore the warning, just rebuild:

```bash
sheetwright build   # source wins, xlsx restored
```

The build-hash record is updated and the warning stops.

## See also

- [CLI reference: import](../reference/cli.md#import) — every flag.
- [MCP reference: re-import tools](../reference/mcp.md#do_reimport_stagexlsx-project-flattenfalse-forcefalse) —
  same flow over MCP.
- [Source format](../reference/source-format.md) — what gets
  rewritten on a merge.
