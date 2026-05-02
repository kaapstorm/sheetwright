"""Render a WorkbookDiff as human-readable text.

Output style: one section per sheet, then named ranges. Each line
prefixed with a marker (`+` added, `-` removed, `*` changed).
"""

from __future__ import annotations

from io import StringIO

from claudesheets.diff.model import SheetDiff, WorkbookDiff


def render(diff: WorkbookDiff) -> str:
    if diff.is_empty():
        return 'no changes\n'

    buf = StringIO()

    for name in diff.sheets_added:
        buf.write(f'+ Sheet added: {name}\n')
    for name in diff.sheets_removed:
        buf.write(f'- Sheet removed: {name}\n')

    for sd in diff.sheets_changed:
        _render_sheet(buf, sd)

    if (
        diff.named_ranges_added
        or diff.named_ranges_removed
        or diff.named_ranges_changed
    ):
        buf.write('\nNamed ranges:\n')
        for nm in diff.named_ranges_added:
            buf.write(f'  + {nm}\n')
        for nm in diff.named_ranges_removed:
            buf.write(f'  - {nm}\n')
        for nrc in diff.named_ranges_changed:
            buf.write(f'  * {nrc.name}: {nrc.old_ref} -> {nrc.new_ref}\n')

    return buf.getvalue()


def _render_sheet(buf: StringIO, sd: SheetDiff) -> None:
    buf.write(f'\nSheet {sd.name}:\n')

    for addr in sd.cells_added:
        buf.write(f'  + {sd.name}!{addr}\n')
    for addr in sd.cells_removed:
        buf.write(f'  - {sd.name}!{addr}\n')
    for cc in sd.cells_changed:
        old = cc.old_formula or repr(cc.old_value)
        new = cc.new_formula or repr(cc.new_value)
        buf.write(f'  * {sd.name}!{cc.addr}: {old} -> {new}\n')

    if sd.frozen_panes_change is not None:
        fpc = sd.frozen_panes_change
        buf.write(f'  * frozen_panes: {fpc.old} -> {fpc.new}\n')
    if sd.print_area_change is not None:
        pac = sd.print_area_change
        buf.write(f'  * print_area: {pac.old} -> {pac.new}\n')

    for col, (old_w, new_w) in sd.column_widths_changed.items():
        buf.write(f'  * column_width[{col}]: {old_w} -> {new_w}\n')

    for addr in sd.comments_added:
        buf.write(f'  + comment {sd.name}!{addr}\n')
    for addr in sd.comments_removed:
        buf.write(f'  - comment {sd.name}!{addr}\n')
    for addr in sd.comments_changed:
        buf.write(f'  * comment {sd.name}!{addr}\n')

    for addr in sd.formats_changed:
        buf.write(f'  * format {sd.name}!{addr}\n')

    if sd.conditional_formats_changed:
        buf.write('  * conditional formatting changed\n')
    if sd.tables_changed:
        buf.write('  * tables changed\n')
