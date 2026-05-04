# Project conventions

## Layout

- `src/sheetwright/` — library + CLI entry points
- `tests/` — pytest tests, mirroring the package layout
- `claude/specs/` — design specifications
- `claude/plans/` — implementation plans (when written)

## Design specs, plans and reviews

| File                  | Path                                            |
|-----------------------|-------------------------------------------------|
| Design specifications | `claude/specs/YYYY-MM-DD_name-of-document.md`   |
| Implementation plans  | `claude/plans/YYYY-MM-DD_name-of-document.md`   |
| Code reviews          | `claude/reviews/YYYY-MM-DD_name-of-document.md` |

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

`python` must be run using `uv run python ...`

## Testing

Use [testsweet](https://github.com/kaapstorm/testsweet). For links to
documentation, see its
[README.md](https://raw.githubusercontent.com/kaapstorm/testsweet/refs/heads/main/README.md).

## Code style

- Match the surrounding code. The project is young; when in doubt,
  prefer the simplest readable option.
- Type hints on public functions and CLI entry points.
- No ad-hoc comments explaining what code does. Comment only when a
  non-obvious *why* matters.
