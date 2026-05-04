"""Compute the structural diff between two Workbook instances."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from sheetwright.diff.model import (
    CellChange,
    FrozenPanesChange,
    NamedRangeChange,
    PrintAreaChange,
    SheetDiff,
    WorkbookDiff,
    _sheet_diff_is_empty,
)
from sheetwright.model.workbook import Sheet, Workbook


def diff_workbooks(old: Workbook, new: Workbook) -> WorkbookDiff:
    """Return the structural diff from `old` to `new`.

    Conventions:
    - `cells_added` lists cells that exist in `new` but not `old`.
    - `cells_removed` lists cells that exist in `old` but not `new`.
    - `cells_changed` carries `old_value`/`new_value` from the
      respective sides.

    All callers must pass arguments in this direction.
    """
    old_sheets = {s.name: s for s in old.sheets}
    new_sheets = {s.name: s for s in new.sheets}

    sheets_added = tuple(sorted(set(new_sheets) - set(old_sheets)))
    sheets_removed = tuple(sorted(set(old_sheets) - set(new_sheets)))
    common_sheet_names = sorted(set(old_sheets) & set(new_sheets))

    sheet_diffs = []
    for name in common_sheet_names:
        sd = _diff_sheet(old_sheets[name], new_sheets[name])
        if not _sheet_diff_is_empty(sd):
            sheet_diffs.append(sd)

    old_nr = {nr.name: nr for nr in old.named_ranges}
    new_nr = {nr.name: nr for nr in new.named_ranges}
    nr_added = tuple(sorted(set(new_nr) - set(old_nr)))
    nr_removed = tuple(sorted(set(old_nr) - set(new_nr)))
    nr_changed = []
    for name in sorted(set(old_nr) & set(new_nr)):
        if old_nr[name].ref != new_nr[name].ref:
            nr_changed.append(
                NamedRangeChange(
                    name=name,
                    old_ref=old_nr[name].ref,
                    new_ref=new_nr[name].ref,
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


def _diff_sheet(old: Sheet, new: Sheet) -> SheetDiff:
    old_cells = old.cells
    new_cells = new.cells
    cells_added = tuple(sorted(set(new_cells) - set(old_cells)))
    cells_removed = tuple(sorted(set(old_cells) - set(new_cells)))
    cells_changed: List[CellChange] = []
    for addr in sorted(set(old_cells) & set(new_cells)):
        co, cn = old_cells[addr], new_cells[addr]
        if co.value != cn.value or co.formula != cn.formula:
            cells_changed.append(
                CellChange(
                    sheet=old.name,
                    addr=addr,
                    old_value=co.value,
                    new_value=cn.value,
                    old_formula=co.formula,
                    new_formula=cn.formula,
                )
            )

    cw_changed: Dict[str, Tuple[Optional[float], Optional[float]]] = {}
    for col in sorted(set(old.column_widths) | set(new.column_widths)):
        old_w = old.column_widths.get(col)
        new_w = new.column_widths.get(col)
        if old_w != new_w:
            cw_changed[col] = (old_w, new_w)

    fp_change = (
        FrozenPanesChange(old=old.frozen_panes, new=new.frozen_panes)
        if old.frozen_panes != new.frozen_panes
        else None
    )
    pa_change = (
        PrintAreaChange(old=old.print_area, new=new.print_area)
        if old.print_area != new.print_area
        else None
    )

    old_cmts = old.comments
    new_cmts = new.comments
    comments_added = tuple(sorted(set(new_cmts) - set(old_cmts)))
    comments_removed = tuple(sorted(set(old_cmts) - set(new_cmts)))
    comments_changed = tuple(
        sorted(
            addr
            for addr in set(old_cmts) & set(new_cmts)
            if old_cmts[addr] != new_cmts[addr]
        )
    )

    formats_changed = _format_diff(old, new)

    cf_changed = list(old.conditional_formats) != list(new.conditional_formats)
    tables_changed = list(old.tables) != list(new.tables)

    return SheetDiff(
        name=old.name,
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


def _format_diff(old: Sheet, new: Sheet) -> Tuple[str, ...]:
    """Return the cell addresses whose effective format differs.

    Format ids may differ between workbooks even when the format
    content is identical, so we resolve through `formats[id]` and
    compare structurally.

    Cells that exist only on one side are not included here; their
    addition/removal is captured by `cells_added`/`cells_removed`,
    which subsumes any format change. This means a "newly added cell
    with non-default format" appears once (as added), not twice.
    """
    addrs: List[str] = []
    common = set(old.cells) & set(new.cells)
    for addr in sorted(common):
        old_cell = old.cells[addr]
        new_cell = new.cells[addr]
        old_fmt = (
            old.formats.get(old_cell.format_id) if old_cell.format_id else None
        )
        new_fmt = (
            new.formats.get(new_cell.format_id) if new_cell.format_id else None
        )
        if old_fmt != new_fmt:
            addrs.append(addr)
    return tuple(addrs)
