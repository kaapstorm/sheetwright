"""Exception hierarchy for sheetwright."""

from typing import ClassVar


class SheetwrightError(Exception):
    """Base for all sheetwright domain errors."""

    code: ClassVar[str] = ''


class ProjectError(SheetwrightError):
    """A project-level error (missing config, malformed structure)."""


class ImportError_(SheetwrightError):
    """Errors during xlsx import."""


class BuildError(SheetwrightError):
    """Errors during build."""


class XlsxTooLargeError(SheetwrightError):
    """Raised when an xlsx file exceeds the configured size limits."""

    code: ClassVar[str] = 'xlsx_too_large'


class BulkInvalidIdentifierError(SheetwrightError):
    """Raised when a bulk cell identifier is not valid."""

    code: ClassVar[str] = 'bulk_invalid_identifier'


class PathOutsideProjectError(SheetwrightError):
    """Raised when a resolved path escapes the project root."""

    code: ClassVar[str] = 'path_outside_project'


class StaleSessionFormatError(SheetwrightError):
    """Raised when a saved session cannot be parsed by the current version."""

    code: ClassVar[str] = 'stale_session_format'
