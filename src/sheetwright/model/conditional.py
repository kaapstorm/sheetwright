"""Conditional formatting rules.

A typed hierarchy mirroring openpyxl's five rule families:
    CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule.

`Sheet.conditional_formats` holds a list of `ConditionalFormat` (the
union alias). User code dispatches with `isinstance` or `match`, and
each subclass exposes typed fields (no stringly-typed param dicts).

`CFStyle` carries the small subset of visual-style attributes we
round-trip for `CellIsRule`/`FormulaRule` (fill colour + font bold/
italic/colour). Full `DifferentialStyle` fidelity is out of scope.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Tuple, Union


@dataclass(frozen=True)
class CFStyle:
    """Lossy subset of openpyxl's DifferentialStyle.

    Carried by CellIsRule and FormulaRule. ColorScaleRule, DataBarRule,
    and IconSetRule paint themselves and don't accept a CFStyle.

    Sufficient for the common econometric uses (fill highlight + bold
    text). Borders, gradient fills, named numFmts, etc. are dropped on
    round-trip. See test_xlsx_conditional_lossiness.py for the
    documented boundary.
    """

    fill_color: Optional[str] = None
    font_bold: bool = False
    font_italic: bool = False
    font_color: Optional[str] = None


@dataclass(frozen=True)
class ConditionalFormatBase:
    """Common fields shared by every rule family.

    `ranges` is a tuple of A1 range strings (frozen). `priority` is
    Excel's rule-evaluation order (lower runs first). `stop_if_true`
    short-circuits subsequent rules on a match.
    """

    ranges: Tuple[str, ...] = ()
    priority: Optional[int] = None
    stop_if_true: bool = False


@dataclass(frozen=True)
class CellIsRule(ConditionalFormatBase):
    """Comparison-against-value rule, e.g. `>0`, `between 1 and 5`.

    `operator` is one of:
        equal, notEqual, greaterThan, greaterThanOrEqual,
        lessThan, lessThanOrEqual, between, notBetween.
    `formula` is a tuple of 1 (most operators) or 2 (between/
    notBetween) cell-formula strings.
    """

    operator: str = 'equal'
    formula: Tuple[str, ...] = ()
    style: Optional[CFStyle] = None


@dataclass(frozen=True)
class FormulaRule(ConditionalFormatBase):
    """Rule that fires when an arbitrary formula evaluates truthy."""

    formula: str = ''
    style: Optional[CFStyle] = None


@dataclass(frozen=True)
class ColorScaleRule(ConditionalFormatBase):
    """Two- or three-stop colour scale.

    A two-stop scale leaves `mid_*` fields as None; a three-stop scale
    sets all three.
    """

    start_type: str = 'min'
    start_value: Optional[str] = None
    start_color: Optional[str] = None
    mid_type: Optional[str] = None
    mid_value: Optional[str] = None
    mid_color: Optional[str] = None
    end_type: str = 'max'
    end_value: Optional[str] = None
    end_color: Optional[str] = None


@dataclass(frozen=True)
class DataBarRule(ConditionalFormatBase):
    start_type: str = 'min'
    start_value: Optional[str] = None
    end_type: str = 'max'
    end_value: Optional[str] = None
    color: str = 'FF638EC6'
    show_value: bool = True


@dataclass(frozen=True)
class IconSetRule(ConditionalFormatBase):
    icon_style: str = '3TrafficLights1'
    type: str = 'percent'
    values: Tuple[str, ...] = ('0', '33', '67')


ConditionalFormat = Union[
    CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule
]
