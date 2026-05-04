"""Data validation rules."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class DataValidation:
    """A single data validation rule applied to one or more ranges.

    The `type` and `operator` strings match Excel's vocabulary
    ("list", "whole", "decimal", "date", "time", "textLength", "custom"
    for type; "between", "notBetween", "equal", "notEqual",
    "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"
    for operator).
    """

    type: str
    ranges: List[str] = field(default_factory=list)
    operator: Optional[str] = None
    formula1: Optional[str] = None
    formula2: Optional[str] = None
    allow_blank: bool = True
