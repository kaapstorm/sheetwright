"""Read an .xlsx file into a Workbook."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any, Optional

import openpyxl
from openpyxl.cell.cell import Cell as XCell

from claudesheets.model.cell import Cell
from claudesheets.model.format import Border, CellFormat, Fill, Font, Side
from claudesheets.model.workbook import NamedRange, Sheet, Workbook


def _normalise_color(c: object) -> Optional[str]:
    """openpyxl colours can be Color objects, theme refs, or hex strings.

    For Plan 1 we only carry hex RGB. Theme colours and indexed colours are
    dropped (None). Alpha-channel hex (8 chars) is normalised to 6-char RGB.
    """
    if c is None:
        return None
    rgb = getattr(c, 'rgb', None) or (c if isinstance(c, str) else None)
    if not isinstance(rgb, str):
        return None
    if len(rgb) == 8:  # AARRGGBB
        return rgb[2:].upper()
    if len(rgb) == 6:
        return rgb.upper()
    return None


def _read_font(c: XCell) -> Optional[Font]:
    f = c.font
    if not f:
        return None
    has_anything = any(
        [
            f.name,
            f.size,
            f.bold,
            f.italic,
            f.underline,
            _normalise_color(f.color),
        ]
    )
    if not has_anything:
        return None
    return Font(
        name=f.name or None,
        size=float(f.size) if f.size is not None else None,
        bold=bool(f.bold),
        italic=bool(f.italic),
        underline=f.underline if f.underline in ('single', 'double') else None,
        color=_normalise_color(f.color),
    )


def _read_fill(c: XCell) -> Optional[Fill]:
    fill = c.fill
    if not fill or fill.fill_type != 'solid':
        return None
    color = _normalise_color(fill.fgColor)
    if not color:
        return None
    return Fill(color=color)


def _read_side(s: Any) -> Optional[Side]:
    if not s or not s.style:
        return None
    return Side(style=s.style, color=_normalise_color(s.color))


def _read_border(c: XCell) -> Optional[Border]:
    b = c.border
    if not b:
        return None
    sides = {
        n: _read_side(getattr(b, n))
        for n in ('left', 'right', 'top', 'bottom')
    }
    if not any(sides.values()):
        return None
    return Border(**sides)


def _cell_format(c: XCell) -> Optional[CellFormat]:
    font = _read_font(c)
    fill = _read_fill(c)
    border = _read_border(c)
    nf = (
        c.number_format
        if c.number_format and c.number_format != 'General'
        else None
    )
    if not any([font, fill, border, nf]):
        return None
    return CellFormat(font=font, fill=fill, border=border, number_format=nf)


def _format_id(fmt: CellFormat) -> str:
    """Stable, content-addressed id for a CellFormat."""
    h = hashlib.sha1(repr(fmt).encode('utf-8')).hexdigest()[:10]
    return f'f-{h}'


def read_xlsx(path: Path) -> Workbook:
    path = Path(path)
    src = openpyxl.load_workbook(path, data_only=False)
    wb = Workbook(name=path.stem)

    for ws in src.worksheets:
        sheet = Sheet(name=ws.title)
        for col_letter, dim in ws.column_dimensions.items():
            if dim.width is not None:
                sheet.column_widths[col_letter] = float(dim.width)
        for row in ws.iter_rows():
            for c in row:
                if not isinstance(c, XCell):
                    # MergedCell — skip; no independent value or format
                    continue
                fmt = _cell_format(c)
                fmt_id = None
                if fmt is not None:
                    fmt_id = _format_id(fmt)
                    sheet.formats[fmt_id] = fmt

                if c.value is None and fmt_id is None:
                    continue
                if isinstance(c.value, str) and c.value.startswith('='):
                    sheet.set(
                        c.coordinate, Cell(formula=c.value, format_id=fmt_id)
                    )
                elif c.value is None:
                    # blank but formatted
                    sheet.set(c.coordinate, Cell(format_id=fmt_id))
                else:
                    # openpyxl stub types c.value more widely than CellValue
                    # (Decimal, RichText, etc.) — these will be narrowed in
                    # later tasks; for now pass through.
                    sheet.set(
                        c.coordinate,
                        Cell(value=c.value, format_id=fmt_id),  # type: ignore[arg-type]
                    )
        wb.sheets.append(sheet)

    for name, defn in src.defined_names.items():
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
