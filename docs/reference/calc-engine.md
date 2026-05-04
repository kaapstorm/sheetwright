# Calc engine

The plugin interface that evaluates a built `.xlsx` and returns
calculated values. The default backend is LibreOffice headless.

## The interface

`claudesheets.calc.CalcEngine` is a one-method ABC:

```python
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict

from claudesheets.model.cell import CellValue


CalcResult = Dict[str, Dict[str, CellValue]]


class CalcEngine(ABC):
    @abstractmethod
    def evaluate(self, xlsx_path: Path) -> CalcResult:
        '''Return calculated values keyed by sheet, then by A1 address.'''
```

`CalcResult` is a nested dict: `result[sheet_name][a1_address]` is
the calculated value. Empty cells should be omitted (callers treat
absent keys as "no value"). Formula cells *must* be present in the
result — `Model.get` raises if a formula cell is missing.

`CellValue` is `str | int | float | bool | datetime | None`.

## Selecting an engine

The engine is chosen per-project in `claudesheets.toml`:

```toml
[build]
calc_engine = "libreoffice"
```

`get_calc_engine(name)` is the registry:

```python
from claudesheets.calc import get_calc_engine

engine = get_calc_engine('libreoffice')
result = engine.evaluate(Path('build/my-model.xlsx'))
```

Currently the only registered name is `libreoffice`. Adding more is
a small addition to `claudesheets.calc.__init__.get_calc_engine` —
or, for one-off use, you can pass a custom engine directly to
`Model(workbook, engine)`.

## `LibreOfficeEngine`

`claudesheets.calc.libreoffice.LibreOfficeEngine` shells out to
`soffice` (LibreOffice headless) to recalculate the file, then reads
the resulting xlsx with openpyxl in `data_only=True` mode to harvest
the cached calculated values.

```python
LibreOfficeEngine(soffice='soffice', timeout=120.0)
```

| Arg | Default | Description |
| --- | --- | --- |
| `soffice` | `'soffice'` | Path to the LibreOffice binary. |
| `timeout` | `120.0` | Seconds before the subprocess is killed. |

### Strategy

For each `evaluate(xlsx)` call:

1. Create a temp dir.
2. Copy/convert the input xlsx via:

    ```
    soffice --headless --calc \
            -env:UserInstallation=file://<tmp>/profile \
            --convert-to xlsx --outdir <tmp>/out <input>
    ```

   The `--convert-to xlsx` step forces LibreOffice to recalculate
   every formula and persist cached values.

3. Read the converted file with openpyxl `data_only=True` and
   collect every non-empty cell.

The dedicated `UserInstallation` profile makes concurrent runs safe
— LibreOffice locks its user profile, so two instances sharing one
profile will block each other. The temp profile lives for the
duration of the `evaluate` call and is removed afterwards.

### Failure modes

`LibreOfficeError` (a `RuntimeError`) is raised on:

- `FileNotFoundError` — `soffice` is not on `$PATH` (or the
  configured path doesn't exist). Install LibreOffice or pass an
  explicit `soffice=` path.
- `subprocess.TimeoutExpired` — the calc didn't finish within the
  timeout. Increase the timeout or simplify the model.
- Non-zero exit code from `soffice` — message includes stderr.
- No xlsx produced in the output directory — usually a sign that
  the input is corrupt or LibreOffice silently bailed.

## The cache

`recalc` and `snapshot` use a content-addressed cache to avoid
re-running the engine when the built xlsx hasn't changed.

- **Key:** SHA-256 of the built xlsx bytes.
- **Path:** `.claudesheets/calc/<sha256>.json`.
- **Format:**

  ```json
  {
    "key": "4f1a9c...",
    "computed_at": "2026-04-30T10:12:00+00:00",
    "result": {
      "Inputs": {"B1": 0.04},
      "Outputs": {"B1": 1040000.0}
    }
  }
  ```

The cache lives under `.claudesheets/`, which is gitignored by
default. To force a recompute, use `claudesheets recalc --force`,
or delete the file:

```bash
rm .claudesheets/calc/<sha256>.json
```

The xlsx writer is deterministic, so a source change that produces
a byte-identical xlsx will hit the same cache entry. That's the
intended behaviour: the calc-engine output is a function of the
xlsx, full stop.

## Adding a backend

To add (say) an Excel-via-COM engine on Windows:

1. Subclass `CalcEngine` and implement `evaluate`.

   ```python
   from pathlib import Path

   from claudesheets.calc.base import CalcEngine, CalcResult


   class ExcelComEngine(CalcEngine):
       def evaluate(self, xlsx_path: Path) -> CalcResult:
           # Open Excel, force recalc, read calculated values
           ...
           return result
   ```

2. Register it in `claudesheets.calc.get_calc_engine`:

   ```python
   def get_calc_engine(name: str) -> CalcEngine:
       if name == 'libreoffice':
           from claudesheets.calc.libreoffice import LibreOfficeEngine
           return LibreOfficeEngine()
       if name == 'excel-com':
           from my_pkg.excel_com import ExcelComEngine
           return ExcelComEngine()
       raise ValueError(f'unknown calc engine: {name!r}')
   ```

3. Reference it in your project's `claudesheets.toml`:

   ```toml
   [build]
   calc_engine = "excel-com"
   ```

For one-off / test use, you can skip registration entirely and pass
your engine directly to `Model`:

```python
from claudesheets.testing import Model

model = Model(workbook, ExcelComEngine())
```

### Contract

Whatever backend you implement must produce a result that satisfies:

- Every formula cell in the workbook is present in the result.
  `Model.get` raises if a formula cell is missing.
- Empty cells should be omitted, not included as `None`.
- Datetime values may be returned as `datetime` objects;
  `snapshot_from_calc_result` will normalise them to ISO-8601
  strings before persisting.
- The function should be pure — same xlsx in, same result out. The
  cache assumes this.

## See also

- [CLI reference: recalc](cli.md#recalc) for cache flags.
- [Testing reference](testing.md) for `Model` and how it consumes
  `CalcResult`.
- [Source format: bulk data](source-format.md#bulk-data-datacsv) for
  the sqlite cache (a separate, build-time cache, not the calc
  cache).
