# Source format

The on-disk format that claudesheets compiles into an `.xlsx`.

A claudesheets project is a directory of plain text. Two TOML
manifests at the root describe the project; one Markdown file per
sheet holds cell values and formulas; an optional YAML sidecar per
sheet carries everything that doesn't fit a table (formats,
validation, conditional formatting, etc.).

## Project layout

```
my-model/
├── claudesheets.toml         # project config
├── workbook.toml             # workbook manifest
├── sheets/
│   ├── 01_inputs.md          # values + formulas
│   ├── 01_inputs.yaml        # formats, validation, etc. (optional)
│   ├── 02_assumptions.md
│   └── 02_outputs.md
├── data/
│   ├── countries.csv         # bulk lookup data
│   └── product_catalog.csv
├── tests/
│   ├── __init__.py
│   ├── test_revenue.py
│   └── snapshots/
│       └── my-model.json     # golden-file snapshot of formula values
├── imports/                  # archived xlsx (with --archive)
│   └── 2026-04-30T10-12-00_model.xlsx
├── build/                    # gitignored build artefacts
│   └── my-model.xlsx
└── .claudesheets/            # gitignored caches
    ├── bulk.sqlite           # cached SQLite of data/*.csv
    ├── build-hash.json       # SHA-256 of last successful build
    ├── reimport.json         # staged re-import session, if any
    └── calc/
        └── <sha256>.json     # cached calc results, keyed by xlsx hash
```

The `init` command writes a default `.gitignore`:

```
build/
.claudesheets/
```

`imports/` is checked in (so reviewers can see what landed when);
everything under `build/` and `.claudesheets/` is reproducible from
the source tree and the calc engine.

## `claudesheets.toml`

Project-level config. Two sections:

```toml
[project]
name = "my-model"

[build]
calc_engine = "libreoffice"
```

| Key | Type | Default | Description |
| --- | --- | --- | --- |
| `project.name` | str | required | Project name. Also the default workbook name. |
| `build.calc_engine` | str | `"libreoffice"` | Calc engine identifier (see [calc engine](calc-engine.md)). |

## `workbook.toml`

Workbook manifest. Lists sheets in workbook order and declares
named ranges.

```toml
[workbook]
name = "my-model"
sheets = ["Inputs", "Assumptions", "Outputs"]

[[named_ranges]]
name = "growth_rate"
scope = "workbook"
ref = "Inputs!$B$1"

[[named_ranges]]
name = "tax_table"
scope = "workbook"
ref = "Assumptions!$A$10:$B$25"
```

| Key | Type | Description |
| --- | --- | --- |
| `workbook.name` | str | Workbook name (used as the built xlsx filename: `build/<name>.xlsx`). |
| `workbook.sheets` | list[str] | Sheet display names in workbook order. |
| `named_ranges[].name` | str | Name as referenced in formulas. |
| `named_ranges[].scope` | str | `"workbook"` (currently only this value). |
| `named_ranges[].sheet` | str (optional) | Sheet for sheet-scoped names (Tier 2). |
| `named_ranges[].ref` | str | Address: `Sheet!$A$1` or `Sheet!$A$1:$B$10`. |

Each sheet name in `workbook.sheets` maps to a file
`sheets/<NN>_<slug>.md`, where `NN` is the 1-based index zero-padded
to two digits and `<slug>` is the sheet name with non-alphanumerics
collapsed to `_` and lowercased. So `["Inputs", "Outputs"]` becomes
`01_inputs.md` and `02_outputs.md`.

Display names live in the manifest; the filename is derived. If you
rename a sheet, the file rename is mechanical — `claudesheets import`
on a re-import or `write_source` will rewrite filenames to match.

## Markdown table format (`sheets/<n>_<slug>.md`)

One table per sheet. The first column is reserved for the row
number; subsequent columns are Excel-style column letters.

```markdown
| (cell) | A             | B                          | C |
| ---    | ---           | ---                        | --- |
| 1      | growth_rate   | 0.04                       |   |
| 2      | base_revenue  | 1000000                    |   |
| 3      | revenue_y1    | =B2*(1+B1)                 |   |
| 4      | active        | TRUE                       |   |
| 5      | label         | Year-over-year (2025→2026) |   |
```

Cell content rules:

- Empty cell: empty string between the bars.
- String: written verbatim. `|` is escaped as `\|` and `\` as `\\`.
- Integer: written as `repr()` (no quotes).
- Float: `repr()` to preserve precision.
- Boolean: `TRUE` or `FALSE` (Excel's convention).
- Formula: starts with `=`. Free-form Excel formula text.

What is **not** in the `.md`:

- Number formats, fonts, fills, borders.
- Column widths.
- Data validation.
- Comments.
- Conditional formatting.
- Tables, frozen panes, print area.

All of those live in the YAML sidecar. The split is intentional:
the `.md` is the human-eye view; the `.yaml` is the cell-decoration
view. Diffs on the `.md` are about *what the model says*; diffs on
the `.yaml` are about *how it's presented*.

## YAML sidecar (`sheets/<n>_<slug>.yaml`)

Optional. Created automatically by `import` and `write_source` when
a sheet has any of the features it covers.

```yaml
column_widths:
  A: 20.0
  B: 14.5

frozen_panes: B2

print_area: A1:H50

formats:
  header:
    font: { name: Calibri, size: 11.0, bold: true }
    fill: { color: "FFCCE5FF" }
    border:
      bottom: { style: medium, color: "FF000000" }
  pct:
    number_format: "0.00%"

cell_formats:
  A1: header
  B1: pct

validations:
  - type: list
    ranges: ["B5:B20"]
    formula1: '"low,medium,high"'
    allow_blank: true

comments:
  B1:
    text: "Set by FP&A; reviewed quarterly."
    author: NH

conditional_formats:
  - kind: cell_is
    ranges: ["B5:B20"]
    operator: greaterThan
    formula: ["0.10"]
    style: { fill_color: "FFFFE0E0", font_bold: true }
  - kind: color_scale
    ranges: ["C5:C20"]
    start_type: min
    start_color: "FFF8696B"
    end_type: max
    end_color: "FF63BE7B"

tables:
  - name: revenue_by_region
    ref: "A10:D40"
    header_row: true
    style: TableStyleMedium2
    columns:
      - { name: region }
      - { name: revenue }
      - { name: cogs }
      - { name: margin, formula: "[@revenue]-[@cogs]" }
```

Top-level keys (all optional):

| Key | Description |
| --- | --- |
| `column_widths` | Map column letter → width (Excel units). |
| `frozen_panes` | A1 address (e.g. `B2`); rows above and columns left are frozen. |
| `print_area` | Excel range (e.g. `A1:H50`). |
| `formats` | Named format definitions. |
| `cell_formats` | Map A1 address → format name. |
| `validations` | List of data validation rules. |
| `comments` | Map A1 address → `{text, author}`. |
| `conditional_formats` | List of conditional-formatting rules. |
| `tables` | List of structured tables (Excel ListObjects). |

### Format definitions

A format is referenced by name from `cell_formats`. Each format may
include `font`, `fill`, `border`, and `number_format`.

```yaml
formats:
  pct:
    number_format: "0.0%"
  bold_red:
    font: { bold: true, color: "FFCC0000" }
    fill: { color: "FFFFEEEE" }
    border:
      top: { style: thin, color: "FF000000" }
```

### Conditional formatting

Five kinds: `cell_is`, `formula`, `color_scale`, `data_bar`,
`icon_set`. See `src/claudesheets/model/conditional.py` for the full
schema. The example above shows the two most common.

### Tables

`tables` declares Excel ListObject (structured-table) ranges with
optional table-style and per-column formulas. A formula on a column
applies to every row in that column.

## Bulk data (`data/*.csv`)

CSVs in `data/` are loaded into `.claudesheets/bulk.sqlite` on every
build. Each file becomes a SQLite table named after the filename
(non-alphanumerics collapsed to `_`, lowercased). The first row is
the header; all columns are typed `TEXT`.

The build is incremental: a CSV that hasn't changed since the last
build (mtime check) is skipped. Otherwise the whole sqlite is
rebuilt deterministically.

The sqlite is intended for tools that prefer SQL to manual lookup
formulas. The current build uses it as a cache; future plans expose
it through a sheet-side `source:` directive.

## Versioning recommendations

Commit:

- `claudesheets.toml`, `workbook.toml`
- `sheets/`
- `data/`
- `tests/` (including `tests/snapshots/`)
- `imports/` (your archived xlsx history)

Don't commit (the default `.gitignore` covers these):

- `build/` — regenerated by `claudesheets build`.
- `.claudesheets/` — calc cache, build hash, bulk sqlite, staged
  re-import sessions.

If your team works directly in Excel sometimes, you might choose to
commit `build/<name>.xlsx` so non-claudesheets users can pull from
git. That's fine; just remember that the canonical source is
`sheets/`, and the `recalc` command will warn if the committed xlsx
diverges from the source-derived build.

## See also

- [CLI reference](cli.md) for the commands that read and write this
  layout.
- [Greenfield tutorial](../tutorials/greenfield-project.md) for an
  end-to-end example.
- [Importing existing](../tutorials/importing-existing.md) for the
  inverse direction (xlsx → source).
