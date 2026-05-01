"""Internal-only pytest fixture: skip when LibreOffice is missing.

Lives under tests/ (not the public package) because it depends on
pytest, while the user-facing test framework is testsweet.
"""

from __future__ import annotations

import shutil

import pytest
from unmagic import fixture


@fixture
def requires_libreoffice():
    """Skip the test if `soffice` is not on $PATH."""
    if shutil.which('soffice') is None:
        pytest.skip('LibreOffice (soffice) not on $PATH')
    yield
