"""Translate domain objects into JSON-serialisable dicts for MCP."""

from __future__ import annotations

from typing import Any, Dict, List

from sheetwright.diff.check import CheckIssue
from sheetwright.diff.format import render
from sheetwright.diff.model import SheetDiff, WorkbookDiff


def diff_to_dict(diff: WorkbookDiff) -> Dict[str, Any]:
    return {
        'is_empty': diff.is_empty(),
        'rendered': render(diff),
        'structured': {
            'sheets_added': list(diff.sheets_added),
            'sheets_removed': list(diff.sheets_removed),
            'sheets_changed': [
                _sheet_diff_to_dict(sd) for sd in diff.sheets_changed
            ],
            'named_ranges_added': list(diff.named_ranges_added),
            'named_ranges_removed': list(diff.named_ranges_removed),
            'named_ranges_changed': [
                {
                    'name': nrc.name,
                    'old_ref': nrc.old_ref,
                    'new_ref': nrc.new_ref,
                }
                for nrc in diff.named_ranges_changed
            ],
        },
    }


def _sheet_diff_to_dict(sd: SheetDiff) -> Dict[str, Any]:
    out: Dict[str, Any] = {
        'name': sd.name,
        'cells_added': list(sd.cells_added),
        'cells_removed': list(sd.cells_removed),
        'cells_changed': [
            {
                'addr': cc.addr,
                'old_value': cc.old_value,
                'new_value': cc.new_value,
                'old_formula': cc.old_formula,
                'new_formula': cc.new_formula,
            }
            for cc in sd.cells_changed
        ],
        'comments_added': list(sd.comments_added),
        'comments_removed': list(sd.comments_removed),
        'comments_changed': list(sd.comments_changed),
        'formats_changed': list(sd.formats_changed),
        'conditional_formats_changed': sd.conditional_formats_changed,
        'tables_changed': sd.tables_changed,
    }
    if sd.frozen_panes_change is not None:
        out['frozen_panes'] = {
            'old': sd.frozen_panes_change.old,
            'new': sd.frozen_panes_change.new,
        }
    if sd.print_area_change is not None:
        out['print_area'] = {
            'old': sd.print_area_change.old,
            'new': sd.print_area_change.new,
        }
    if sd.column_widths_changed:
        out['column_widths'] = {
            col: {'old': old, 'new': new}
            for col, (old, new) in sd.column_widths_changed.items()
        }
    return out


def check_issues_to_dicts(
    issues: List[CheckIssue],
) -> List[Dict[str, Any]]:
    return [
        {
            'kind': issue.kind,
            'detail': issue.detail,
            'location': issue.location,
        }
        for issue in issues
    ]
