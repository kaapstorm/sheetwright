"""Write a Workbook to an .xlsx file."""

from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from claudesheets.model.workbook import Workbook


def write_xlsx(wb: Workbook, path: Path) -> None:
    path = Path(path)
    out = openpyxl.Workbook()
    # openpyxl creates a default "Sheet"; remove it.
    default = out.active
    if default is not None:
        out.remove(default)

    for sheet in wb.sheets:
        ws = out.create_sheet(title=sheet.name)
        for addr, cell in sheet.cells.items():
            if cell.formula is not None:
                ws[addr] = cell.formula
            else:
                ws[addr] = cell.value

    for nr in wb.named_ranges:
        defn = DefinedName(name=nr.name, attr_text=nr.ref)
        if nr.scope == 'workbook':
            out.defined_names[nr.name] = defn
        else:
            assert nr.sheet is not None
            out[nr.sheet].defined_names[nr.name] = defn

    out.save(path)
