"""Skip decorator for libreoffice-gated tests."""

from __future__ import annotations

import shutil

from testsweet import skip


requires_libreoffice = skip(
    condition=lambda: shutil.which('soffice') is None,
    reason='LibreOffice (soffice) not on $PATH',
)
