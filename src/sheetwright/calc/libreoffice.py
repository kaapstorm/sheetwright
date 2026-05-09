"""LibreOffice headless calc engine.

Strategy: copy the input xlsx into a fresh temp dir, run
`soffice --headless --calc --convert-to xlsx --outdir <tmp> <input>`
which forces LibreOffice to recalculate every formula and persist
cached values, then load the converted file with openpyxl in
`data_only=True` mode and read every cell.

We use `-env:UserInstallation=file://<tmp>/profile` to give each run
its own profile directory so concurrent invocations do not block each
other on LibreOffice's user-profile lock.
"""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path
from typing import Dict, cast

from sheetwright.calc.base import CalcEngine, CalcResult
from sheetwright.model.cell import CellValue
from sheetwright.security import SecurityLimits
from sheetwright.xlsx.safe_load import safe_load_workbook


class LibreOfficeError(RuntimeError):
    """soffice exited non-zero or produced no output."""


class LibreOfficeEngine(CalcEngine):
    def __init__(self, soffice: str = 'soffice'):
        self.soffice = soffice

    def evaluate(
        self, xlsx_path: Path, *, limits: SecurityLimits
    ) -> CalcResult:
        xlsx_path = Path(xlsx_path).resolve()
        with tempfile.TemporaryDirectory(prefix='sheetwright-calc-') as td_str:
            td = Path(td_str)
            profile = td / 'profile'
            outdir = td / 'out'
            outdir.mkdir()

            cmd = [
                self.soffice,
                '--headless',
                '--safe-mode',
                '--norestore',
                '--nolockcheck',
                '--nofirststartwizard',
                '--nodefault',
                '--calc',
                f'-env:UserInstallation={profile.as_uri()}',
                '--convert-to',
                'xlsx',
                '--outdir',
                str(outdir),
                str(xlsx_path),
            ]
            timeout = limits.soffice_timeout
            try:
                proc = subprocess.run(
                    cmd,
                    capture_output=True,
                    timeout=timeout,
                    check=False,
                )
            except FileNotFoundError as e:
                raise LibreOfficeError(
                    f'{self.soffice}: not on $PATH or not executable'
                ) from e
            except subprocess.TimeoutExpired as e:
                raise LibreOfficeError(
                    f'{self.soffice} timed out after {timeout}s'
                ) from e

            if proc.returncode != 0:
                stderr = proc.stderr.decode('utf-8', 'replace').strip()
                stdout = proc.stdout.decode('utf-8', 'replace').strip()
                detail = stderr or stdout or '(no output)'
                raise LibreOfficeError(
                    f'soffice exited {proc.returncode}: {detail}'
                )

            converted = outdir / xlsx_path.name
            if not converted.is_file():
                candidates = list(outdir.glob('*.xlsx'))
                if len(candidates) != 1:
                    raise LibreOfficeError(
                        f'expected 1 xlsx in {outdir}, found {candidates}'
                    )
                converted = candidates[0]

            return _read_calculated(converted, limits)


def _read_calculated(xlsx_path: Path, limits: SecurityLimits) -> CalcResult:
    wb = safe_load_workbook(xlsx_path, limits, data_only=True, read_only=True)
    out: CalcResult = {}
    for ws in wb.worksheets:
        sheet: Dict[str, CellValue] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                sheet[cell.coordinate] = cast(CellValue, cell.value)
        out[ws.title] = sheet
    return out
