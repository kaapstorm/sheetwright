"""Exception hierarchy for claudesheets."""


class ClaudesheetsError(Exception):
    """Base class for all claudesheets errors."""


class ProjectError(ClaudesheetsError):
    """A project-level error (missing config, malformed structure)."""


class ImportError_(ClaudesheetsError):
    """Errors during xlsx import."""


class BuildError(ClaudesheetsError):
    """Errors during build."""
