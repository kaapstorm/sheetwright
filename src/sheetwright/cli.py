"""sheetwright CLI entry point."""

from __future__ import annotations

import sys
from typing import Optional

import click

from sheetwright.security import OperatorLimits

_operator_limits: Optional[OperatorLimits] = None


def get_operator_limits() -> OperatorLimits:
    """Return the process-wide operator limits (built once at CLI entry)."""
    assert _operator_limits is not None, 'CLI not initialised'
    return _operator_limits


@click.group(context_settings={'help_option_names': ['-h', '--help']})
@click.version_option(package_name='sheetwright')
def main() -> None:
    """Work with spreadsheets from Claude Code."""
    global _operator_limits
    _operator_limits = OperatorLimits.from_environment()


@main.command('init')
@click.argument('path', type=click.Path(file_okay=False), default='.')
def init_cmd(path: str) -> None:
    """Scaffold an empty sheetwright project."""
    from sheetwright.commands.init_cmd import run

    run(path)


@main.command('import')
@click.argument(
    'xlsx',
    type=click.Path(exists=True, dir_okay=False),
    required=False,
)
@click.option('--archive', is_flag=True, help='Copy the xlsx into imports/.')
@click.option(
    '--flatten',
    is_flag=True,
    help='Replace external-reference formulas with cached values.',
)
@click.option(
    '-I',
    '--non-interactive',
    'non_interactive',
    is_flag=True,
    help='Print the diff and exit; use --apply or --abort to follow up.',
)
@click.option(
    '--apply',
    is_flag=True,
    help='Apply a previously staged re-import.',
)
@click.option(
    '--abort',
    is_flag=True,
    help='Discard a previously staged re-import.',
)
@click.option(
    '--force',
    is_flag=True,
    help='Skip the uncommitted-source guard during re-import.',
)
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the sheetwright project.',
)
def import_cmd(
    xlsx: Optional[str],
    archive: bool,
    flatten: bool,
    non_interactive: bool,
    apply: bool,
    abort: bool,
    force: bool,
    project_path: str,
) -> None:
    """Read an .xlsx file into source form, or merge updates into an
    existing project."""
    from sheetwright.commands.import_cmd import run

    run(
        xlsx_path=xlsx,
        project_path=project_path,
        archive=archive,
        flatten=flatten,
        non_interactive=non_interactive,
        apply=apply,
        abort=abort,
        force=force,
    )


@main.command('build')
@click.option(
    '--out',
    'out_path',
    type=click.Path(dir_okay=False),
    default=None,
    help='Output path. Defaults to build/<workbook-name>.xlsx.',
)
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the sheetwright project.',
)
def build_cmd(out_path: str | None, project_path: str) -> None:
    """Compile source files into an .xlsx."""
    from sheetwright.commands.build_cmd import run

    run(project_path=project_path, out_path=out_path)


@main.command('recalc')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the sheetwright project.',
)
@click.option(
    '--force',
    is_flag=True,
    help='Ignore the cache and re-run the calc engine.',
)
def recalc_cmd(project_path: str, force: bool) -> None:
    """Run the calc engine and cache calculated values."""
    from sheetwright.commands.recalc_cmd import run

    run(project_path=project_path, force=force)


@main.command('snapshot')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the sheetwright project.',
)
@click.option(
    '--update',
    is_flag=True,
    help='Overwrite the saved snapshot with the current calculated values.',
)
def snapshot_cmd(project_path: str, update: bool) -> None:
    """Compare or update the golden-file snapshot of calculated values."""
    from sheetwright.commands.snapshot_cmd import run

    run(project_path=project_path, update=update)


@main.command('diff')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the sheetwright project.',
)
@click.option(
    '--vs',
    'vs',
    type=str,
    default=None,
    help='Comparison target: "xlsx:<path>" or "source:<path>". '
    'Default: build/<name>.xlsx of this project.',
)
def diff_cmd(project_path: str, vs: Optional[str]) -> None:
    """Show a semantic diff between source and a target workbook."""
    from sheetwright.commands.diff_cmd import run

    run(project_path=project_path, vs=vs)


@main.command('check')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the sheetwright project.',
)
def check_cmd(project_path: str) -> None:
    """Lint dangling refs, missing names, schema mismatches."""
    from sheetwright.commands.check_cmd import run

    run(project_path=project_path)


@main.command(
    'test',
    context_settings={'ignore_unknown_options': True},
)
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the sheetwright project.',
)
@click.argument('targets', nargs=-1, type=click.UNPROCESSED)
def test_cmd(project_path: str, targets: tuple[str, ...]) -> None:
    """Run the project's testsweet tests."""
    from sheetwright.commands.test_cmd import run

    run(project_path=project_path, targets=list(targets))


@main.command('mcp')
def mcp_cmd() -> None:
    """Run the MCP server on stdio.

    Most MCP clients launch this subcommand as a subprocess and
    communicate via stdin/stdout. The server stays alive until the
    client closes the connection.
    """
    from sheetwright.commands.mcp_cmd import run

    run()


if __name__ == '__main__':  # pragma: no cover
    sys.exit(main())
