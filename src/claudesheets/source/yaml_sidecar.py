"""Read/write the per-sheet YAML sidecar.

The sidecar carries everything that doesn't fit cleanly in a Markdown
table: column widths, format definitions, the cell -> format mapping,
and data validation rules.

Cell values and formulas live in the .md, not here.
"""

from __future__ import annotations

import io
from dataclasses import asdict
from typing import Any, Dict, Optional

from ruamel.yaml import YAML

from claudesheets.model.cell import Cell
from claudesheets.model.format import Border, CellFormat, Fill, Font, Side
from claudesheets.model.validation import DataValidation
from claudesheets.model.workbook import Sheet

_yaml = YAML(typ='rt')
_yaml.indent(mapping=2, sequence=4, offset=2)
_yaml.preserve_quotes = True


def _font_to_dict(f: Font) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for k in ('name', 'size', 'bold', 'italic', 'underline', 'color'):
        v = getattr(f, k)
        if v not in (None, False):
            out[k] = v
    return out


def _font_from_dict(d: Optional[Dict[str, Any]]) -> Optional[Font]:
    if not d:
        return None
    return Font(
        name=d.get('name'),
        size=d.get('size'),
        bold=bool(d.get('bold', False)),
        italic=bool(d.get('italic', False)),
        underline=d.get('underline'),
        color=d.get('color'),
    )


def _side_to_dict(s: Optional[Side]) -> Optional[Dict[str, Any]]:
    if s is None:
        return None
    out: Dict[str, Any] = {'style': s.style}
    if s.color:
        out['color'] = s.color
    return out


def _side_from_dict(d: Optional[Dict[str, Any]]) -> Optional[Side]:
    if not d:
        return None
    return Side(style=d.get('style'), color=d.get('color'))


def _fmt_to_dict(fmt: CellFormat) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    if fmt.font is not None:
        out['font'] = _font_to_dict(fmt.font)
    if fmt.fill is not None and fmt.fill.color:
        out['fill'] = {'color': fmt.fill.color}
    if fmt.border is not None:
        b: Dict[str, Any] = {}
        for n in ('left', 'right', 'top', 'bottom'):
            sd = _side_to_dict(getattr(fmt.border, n))
            if sd is not None:
                b[n] = sd
        if b:
            out['border'] = b
    if fmt.number_format:
        out['number_format'] = fmt.number_format
    return out


def _fmt_from_dict(d: Dict[str, Any]) -> CellFormat:
    border_d = d.get('border')
    border = None
    if border_d:
        border = Border(
            left=_side_from_dict(border_d.get('left')),
            right=_side_from_dict(border_d.get('right')),
            top=_side_from_dict(border_d.get('top')),
            bottom=_side_from_dict(border_d.get('bottom')),
        )
    fill_d = d.get('fill')
    fill = (
        Fill(color=fill_d.get('color'))
        if fill_d and fill_d.get('color')
        else None
    )
    return CellFormat(
        font=_font_from_dict(d.get('font')),
        fill=fill,
        border=border,
        number_format=d.get('number_format'),
    )


def dump_yaml(sheet: Sheet) -> str:
    doc: Dict[str, Any] = {}

    if sheet.column_widths:
        doc['column_widths'] = dict(sheet.column_widths)

    if sheet.frozen_panes:
        doc['frozen_panes'] = sheet.frozen_panes

    if sheet.formats:
        doc['formats'] = {
            fid: _fmt_to_dict(f) for fid, f in sheet.formats.items()
        }

    cell_formats = {
        a: c.format_id for a, c in sheet.cells.items() if c.format_id
    }
    if cell_formats:
        doc['cell_formats'] = cell_formats

    if sheet.validations:
        doc['validations'] = [
            {
                k: v
                for k, v in asdict(dv).items()
                if v not in (None, [], False) or k == 'ranges'
            }
            for dv in sheet.validations
        ]

    buf = io.StringIO()
    _yaml.dump(doc, buf)
    return buf.getvalue()


def load_yaml(sheet: Sheet, text: str) -> None:
    if not text.strip():
        return
    doc = _yaml.load(text) or {}

    for col, w in (doc.get('column_widths') or {}).items():
        sheet.column_widths[col] = float(w)

    if doc.get('frozen_panes'):
        sheet.frozen_panes = str(doc['frozen_panes'])

    for fid, d in (doc.get('formats') or {}).items():
        sheet.formats[fid] = _fmt_from_dict(dict(d))

    for addr, fid in (doc.get('cell_formats') or {}).items():
        existing = sheet.get(addr)
        sheet.set(
            addr,
            Cell(
                value=existing.value,
                formula=existing.formula,
                format_id=fid,
            ),
        )

    for d in doc.get('validations') or []:
        sheet.validations.append(
            DataValidation(
                type=d['type'],
                ranges=list(d.get('ranges') or []),
                operator=d.get('operator'),
                formula1=d.get('formula1'),
                formula2=d.get('formula2'),
                allow_blank=bool(d.get('allow_blank', True)),
            )
        )
