"""Write a Workbook to an .xlsx file."""

from __future__ import annotations

from pathlib import Path
from typing import Literal, Optional

import openpyxl
from openpyxl.cell.cell import Cell as XCell
from openpyxl.worksheet.datavalidation import DataValidation as XDV
from openpyxl.styles import (
    Border as XBorder,
    Color,
    Font as XFont,
    PatternFill,
    Side as XSide,
)
from openpyxl.workbook.defined_name import DefinedName

from claudesheets.model.format import Border, CellFormat, Fill, Font, Side
from claudesheets.model.workbook import Workbook

# openpyxl's underline Literal type
_Underline = Optional[
    Literal['single', 'double', 'singleAccounting', 'doubleAccounting', 'none']
]


def _xfont(font: Optional[Font]) -> Optional[XFont]:
    if font is None:
        return None
    color = Color(rgb=('FF' + font.color)) if font.color else None
    underline: _Underline = font.underline  # type: ignore[assignment]
    return XFont(
        name=font.name,
        size=font.size,
        bold=font.bold,
        italic=font.italic,
        underline=underline,
        color=color,
    )


def _xfill(fill: Optional[Fill]) -> Optional[PatternFill]:
    if fill is None or not fill.color:
        return None
    return PatternFill(
        fill_type='solid', fgColor=Color(rgb=('FF' + fill.color))
    )


def _xside(side: Optional[Side]) -> Optional[XSide]:
    if side is None or not side.style:
        return None
    color = Color(rgb=('FF' + side.color)) if side.color else None
    return XSide(style=side.style, color=color)  # type: ignore[arg-type]


def _xborder(border: Optional[Border]) -> Optional[XBorder]:
    if border is None:
        return None
    return XBorder(
        left=_xside(border.left),
        right=_xside(border.right),
        top=_xside(border.top),
        bottom=_xside(border.bottom),
    )


def _apply_format(cell: XCell, fmt: CellFormat) -> None:
    if fmt.font is not None:
        cell.font = _xfont(fmt.font)  # type: ignore[assignment]
    if fmt.fill is not None:
        cell.fill = _xfill(fmt.fill)  # type: ignore[assignment]
    if fmt.border is not None:
        cell.border = _xborder(fmt.border)  # type: ignore[assignment]
    if fmt.number_format is not None:
        cell.number_format = fmt.number_format


def write_xlsx(wb: Workbook, path: Path) -> None:
    path = Path(path)
    out = openpyxl.Workbook()
    default = out.active
    if default is not None:
        out.remove(default)

    for sheet in wb.sheets:
        ws = out.create_sheet(title=sheet.name)

        for col, width in sheet.column_widths.items():
            ws.column_dimensions[col].width = width

        for v in sheet.validations:
            xdv = XDV(
                type=v.type,  # type: ignore[arg-type]
                operator=v.operator,  # type: ignore[arg-type]
                formula1=v.formula1,
                formula2=v.formula2,
                allow_blank=v.allow_blank,
            )
            for r in v.ranges:
                xdv.add(r)
            ws.add_data_validation(xdv)

        for addr, cell in sheet.cells.items():
            xc = ws[addr]
            if cell.formula is not None:
                xc.value = cell.formula
            elif cell.value is not None:
                xc.value = cell.value
            if cell.format_id and cell.format_id in sheet.formats:
                _apply_format(xc, sheet.formats[cell.format_id])

    for nr in wb.named_ranges:
        defn = DefinedName(name=nr.name, attr_text=nr.ref)
        if nr.scope == 'workbook':
            out.defined_names[nr.name] = defn
        else:
            assert nr.sheet is not None
            out[nr.sheet].defined_names[nr.name] = defn

    out.save(path)
