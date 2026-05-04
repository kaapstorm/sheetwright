"""Exception hierarchy for sheetwright."""


class SheetwrightError(Exception):
    """Base class for all sheetwright errors."""


class ProjectError(SheetwrightError):
    """A project-level error (missing config, malformed structure)."""


class ImportError_(SheetwrightError):
    """Errors during xlsx import."""


class BuildError(SheetwrightError):
    """Errors during build."""
