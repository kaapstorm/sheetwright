"""Build a small xlsx with an external reference, plus a cached value.

Pure openpyxl — we craft the formula and bake the cached value via
`cell.value = number` after writing the formula via `cell._value`-
style, since openpyxl's Cell doesn't naturally let us set both
formula and cached value. The simplest path is to write the workbook
once, then patch the underlying xml.

For tests, we go simpler: write a regular formula referencing
`'[other.xlsx]Sheet1'!A1` and rely on the fact that openpyxl will
expose this formula's cached value via `data_only=True` if the file
has one. We seed a cached value by manually patching the saved xml.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import openpyxl


def write_xlsx_with_external_ref(
    path: Path, cached_value: float = 42.0
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'S'
    ws['A1'] = "='[other.xlsx]Sheet1'!$A$1"
    wb.save(path)

    # Inject a cached value into the formula cell so data_only=True
    # produces something. The xml lives in xl/worksheets/sheet1.xml.
    with zipfile.ZipFile(path, 'r') as z:
        files = {n: z.read(n) for n in z.namelist()}

    sheet_xml = files['xl/worksheets/sheet1.xml'].decode('utf-8')
    sheet_xml = re.sub(
        r'(<f[^>]*>[^<]*</f>)',
        r'\1<v>' + str(cached_value) + r'</v>',
        sheet_xml,
        count=1,
    )
    files['xl/worksheets/sheet1.xml'] = sheet_xml.encode('utf-8')

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
