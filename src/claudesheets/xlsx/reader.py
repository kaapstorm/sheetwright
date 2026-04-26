"""Read an .xlsx file into a Workbook."""

from __future__ import annotations

from pathlib import Path

import openpyxl

from claudesheets.model.cell import Cell
from claudesheets.model.workbook import NamedRange, Sheet, Workbook


def read_xlsx(path: Path) -> Workbook:
    path = Path(path)
    src = openpyxl.load_workbook(path, data_only=False)

    wb = Workbook(name=path.stem)

    for ws in src.worksheets:
        sheet = Sheet(name=ws.title)
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if isinstance(c.value, str) and c.value.startswith('='):
                    sheet.set(c.coordinate, Cell(formula=c.value))
                else:
                    # Cast: openpyxl's value type is broader than CellValue;
                    # unsupported types (Decimal, RichText, etc.) are
                    # preserved as-is for now and narrowed in later tasks.
                    sheet.set(
                        c.coordinate,
                        Cell(value=c.value),  # type: ignore[arg-type]
                    )
        wb.sheets.append(sheet)

    for name, defn in src.defined_names.items():
        # openpyxl exposes workbook-scope names via wb.defined_names and
        # sheet-scope names via ws.defined_names. We handle both.
        wb.named_ranges.append(
            NamedRange(name=name, ref=defn.attr_text, scope='workbook')
        )

    for ws in src.worksheets:
        for name, defn in ws.defined_names.items():
            wb.named_ranges.append(
                NamedRange(
                    name=name,
                    ref=defn.attr_text,
                    scope='sheet',
                    sheet=ws.title,
                )
            )

    return wb
