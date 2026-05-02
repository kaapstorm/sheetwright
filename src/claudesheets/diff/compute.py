"""Compute the structural diff between two Workbook instances."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from claudesheets.diff.model import (
    CellChange,
    FrozenPanesChange,
    NamedRangeChange,
    PrintAreaChange,
    SheetDiff,
    WorkbookDiff,
    _sheet_diff_is_empty,
)
from claudesheets.model.workbook import Sheet, Workbook


def diff_workbooks(a: Workbook, b: Workbook) -> WorkbookDiff:
    a_sheets = {s.name: s for s in a.sheets}
    b_sheets = {s.name: s for s in b.sheets}

    sheets_added = tuple(sorted(set(b_sheets) - set(a_sheets)))
    sheets_removed = tuple(sorted(set(a_sheets) - set(b_sheets)))
    common_sheet_names = sorted(set(a_sheets) & set(b_sheets))

    sheet_diffs = []
    for name in common_sheet_names:
        sd = _diff_sheet(a_sheets[name], b_sheets[name])
        if not _sheet_diff_is_empty(sd):
            sheet_diffs.append(sd)

    a_nr = {nr.name: nr for nr in a.named_ranges}
    b_nr = {nr.name: nr for nr in b.named_ranges}
    nr_added = tuple(sorted(set(b_nr) - set(a_nr)))
    nr_removed = tuple(sorted(set(a_nr) - set(b_nr)))
    nr_changed = []
    for name in sorted(set(a_nr) & set(b_nr)):
        if a_nr[name].ref != b_nr[name].ref:
            nr_changed.append(
                NamedRangeChange(
                    name=name,
                    old_ref=a_nr[name].ref,
                    new_ref=b_nr[name].ref,
                )
            )

    return WorkbookDiff(
        sheets_added=sheets_added,
        sheets_removed=sheets_removed,
        sheets_changed=tuple(sheet_diffs),
        named_ranges_added=nr_added,
        named_ranges_removed=nr_removed,
        named_ranges_changed=tuple(nr_changed),
    )


def _diff_sheet(a: Sheet, b: Sheet) -> SheetDiff:
    a_cells = a.cells
    b_cells = b.cells
    cells_added = tuple(sorted(set(b_cells) - set(a_cells)))
    cells_removed = tuple(sorted(set(a_cells) - set(b_cells)))
    cells_changed: List[CellChange] = []
    for addr in sorted(set(a_cells) & set(b_cells)):
        ca, cb = a_cells[addr], b_cells[addr]
        if ca.value != cb.value or ca.formula != cb.formula:
            cells_changed.append(
                CellChange(
                    sheet=a.name,
                    addr=addr,
                    old_value=ca.value,
                    new_value=cb.value,
                    old_formula=ca.formula,
                    new_formula=cb.formula,
                )
            )

    cw_changed: Dict[str, Tuple[Optional[float], Optional[float]]] = {}
    for col in sorted(set(a.column_widths) | set(b.column_widths)):
        old = a.column_widths.get(col)
        new = b.column_widths.get(col)
        if old != new:
            cw_changed[col] = (old, new)

    fp_change = (
        FrozenPanesChange(old=a.frozen_panes, new=b.frozen_panes)
        if a.frozen_panes != b.frozen_panes
        else None
    )
    pa_change = (
        PrintAreaChange(old=a.print_area, new=b.print_area)
        if a.print_area != b.print_area
        else None
    )

    a_cmts = a.comments
    b_cmts = b.comments
    comments_added = tuple(sorted(set(b_cmts) - set(a_cmts)))
    comments_removed = tuple(sorted(set(a_cmts) - set(b_cmts)))
    comments_changed = tuple(
        sorted(
            addr
            for addr in set(a_cmts) & set(b_cmts)
            if a_cmts[addr] != b_cmts[addr]
        )
    )

    formats_changed = _format_diff(a, b)

    cf_changed = list(a.conditional_formats) != list(b.conditional_formats)
    tables_changed = list(a.tables) != list(b.tables)

    return SheetDiff(
        name=a.name,
        cells_added=cells_added,
        cells_removed=cells_removed,
        cells_changed=tuple(cells_changed),
        column_widths_changed=cw_changed,
        frozen_panes_change=fp_change,
        print_area_change=pa_change,
        comments_added=comments_added,
        comments_removed=comments_removed,
        comments_changed=comments_changed,
        conditional_formats_changed=cf_changed,
        tables_changed=tables_changed,
        formats_changed=formats_changed,
    )


def _format_diff(a: Sheet, b: Sheet) -> Tuple[str, ...]:
    """Return the cell addresses whose effective format differs.

    Format ids may differ between workbooks even when the format
    content is identical, so we resolve through `formats[id]` and
    compare structurally.
    """
    addrs: List[str] = []
    common = set(a.cells) & set(b.cells)
    for addr in sorted(common):
        a_cell = a.cells[addr]
        b_cell = b.cells[addr]
        a_fmt = a.formats.get(a_cell.format_id) if a_cell.format_id else None
        b_fmt = b.formats.get(b_cell.format_id) if b_cell.format_id else None
        if a_fmt != b_fmt:
            addrs.append(addr)
    return tuple(addrs)
