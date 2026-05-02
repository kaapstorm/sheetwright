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
from claudesheets.model.comment import Comment as Cmt
from claudesheets.model.conditional import (
    CellIsRule,
    CFStyle,
    ColorScaleRule,
    ConditionalFormat,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
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


def _cfstyle_to_dict(s: CFStyle) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    if s.fill_color:
        out['fill_color'] = s.fill_color
    if s.font_bold:
        out['font_bold'] = True
    if s.font_italic:
        out['font_italic'] = True
    if s.font_color:
        out['font_color'] = s.font_color
    return out


def _cfstyle_from_dict(
    d: Optional[Dict[str, Any]],
) -> Optional[CFStyle]:
    if not d:
        return None
    return CFStyle(
        fill_color=d.get('fill_color'),
        font_bold=bool(d.get('font_bold', False)),
        font_italic=bool(d.get('font_italic', False)),
        font_color=d.get('font_color'),
    )


def _cf_to_dict(cf: ConditionalFormat) -> Dict[str, Any]:
    out: Dict[str, Any] = {'ranges': list(cf.ranges)}
    if cf.priority is not None:
        out['priority'] = cf.priority
    if cf.stop_if_true:
        out['stop_if_true'] = True

    if isinstance(cf, CellIsRule):
        out['kind'] = 'cell_is'
        out['operator'] = cf.operator
        out['formula'] = list(cf.formula)
        if cf.style is not None:
            out['style'] = _cfstyle_to_dict(cf.style)
    elif isinstance(cf, FormulaRule):
        out['kind'] = 'formula'
        out['formula'] = cf.formula
        if cf.style is not None:
            out['style'] = _cfstyle_to_dict(cf.style)
    elif isinstance(cf, ColorScaleRule):
        out['kind'] = 'color_scale'
        out['start_type'] = cf.start_type
        out['start_color'] = cf.start_color
        out['end_type'] = cf.end_type
        out['end_color'] = cf.end_color
        if cf.start_value is not None:
            out['start_value'] = cf.start_value
        if cf.mid_type is not None:
            out['mid_type'] = cf.mid_type
            out['mid_value'] = cf.mid_value
            out['mid_color'] = cf.mid_color
        if cf.end_value is not None:
            out['end_value'] = cf.end_value
    elif isinstance(cf, DataBarRule):
        out['kind'] = 'data_bar'
        out['start_type'] = cf.start_type
        out['end_type'] = cf.end_type
        out['color'] = cf.color
        if cf.start_value is not None:
            out['start_value'] = cf.start_value
        if cf.end_value is not None:
            out['end_value'] = cf.end_value
        out['show_value'] = cf.show_value
    elif isinstance(cf, IconSetRule):
        out['kind'] = 'icon_set'
        out['icon_style'] = cf.icon_style
        out['type'] = cf.type
        out['values'] = list(cf.values)
    else:
        raise TypeError(f'unknown ConditionalFormat type: {type(cf).__name__}')
    return out


def _cf_from_dict(d: Dict[str, Any]) -> ConditionalFormat:
    kind = d.get('kind', 'cell_is')
    common = {
        'ranges': tuple(d.get('ranges') or ()),
        'priority': d.get('priority'),
        'stop_if_true': bool(d.get('stop_if_true', False)),
    }
    if kind == 'cell_is':
        return CellIsRule(
            **common,  # type: ignore[arg-type]
            operator=str(d.get('operator', 'equal')),
            formula=tuple(d.get('formula') or ()),
            style=_cfstyle_from_dict(d.get('style')),
        )
    if kind == 'formula':
        return FormulaRule(
            **common,  # type: ignore[arg-type]
            formula=str(d.get('formula', '')),
            style=_cfstyle_from_dict(d.get('style')),
        )
    if kind == 'color_scale':
        return ColorScaleRule(
            **common,  # type: ignore[arg-type]
            start_type=str(d.get('start_type', 'min')),
            start_value=d.get('start_value'),
            start_color=str(d.get('start_color', 'FFFFFFFF')),
            mid_type=d.get('mid_type'),
            mid_value=d.get('mid_value'),
            mid_color=d.get('mid_color'),
            end_type=str(d.get('end_type', 'max')),
            end_value=d.get('end_value'),
            end_color=str(d.get('end_color', 'FF000000')),
        )
    if kind == 'data_bar':
        return DataBarRule(
            **common,  # type: ignore[arg-type]
            start_type=str(d.get('start_type', 'min')),
            start_value=d.get('start_value'),
            end_type=str(d.get('end_type', 'max')),
            end_value=d.get('end_value'),
            color=str(d.get('color', 'FF638EC6')),
            show_value=bool(d.get('show_value', True)),
        )
    if kind == 'icon_set':
        return IconSetRule(
            **common,  # type: ignore[arg-type]
            icon_style=str(d.get('icon_style', '3TrafficLights1')),
            type=str(d.get('type', 'percent')),
            values=tuple(d.get('values') or ()),
        )
    raise ValueError(f'unknown CF kind in YAML: {kind!r}')


def dump_yaml(sheet: Sheet) -> str:
    doc: Dict[str, Any] = {}

    if sheet.column_widths:
        doc['column_widths'] = dict(sheet.column_widths)

    if sheet.frozen_panes:
        doc['frozen_panes'] = sheet.frozen_panes

    if sheet.print_area:
        doc['print_area'] = sheet.print_area

    if sheet.formats:
        doc['formats'] = {
            fid: _fmt_to_dict(f) for fid, f in sheet.formats.items()
        }

    cell_formats = {
        a: c.format_id for a, c in sheet.cells.items() if c.format_id
    }
    if cell_formats:
        doc['cell_formats'] = cell_formats

    if sheet.comments:
        doc['comments'] = {
            addr: {'author': c.author, 'text': c.text}
            for addr, c in sheet.comments.items()
        }

    if sheet.conditional_formats:
        doc['conditional_formats'] = [
            _cf_to_dict(cf) for cf in sheet.conditional_formats
        ]

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

    if doc.get('print_area'):
        sheet.print_area = str(doc['print_area'])

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

    for addr, d in (doc.get('comments') or {}).items():
        sheet.comments[addr] = Cmt(
            author=str(d.get('author', '')),
            text=str(d.get('text', '')),
        )

    for d in doc.get('conditional_formats') or []:
        sheet.conditional_formats.append(_cf_from_dict(dict(d)))

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
