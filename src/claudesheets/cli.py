"""claudesheets CLI entry point."""

from __future__ import annotations

import sys
from typing import Optional

import click


@click.group(context_settings={'help_option_names': ['-h', '--help']})
@click.version_option(package_name='claudesheets')
def main() -> None:
    """Work with spreadsheets from Claude Code."""


@main.command('init')
@click.argument('path', type=click.Path(file_okay=False), default='.')
def init_cmd(path: str) -> None:
    """Scaffold an empty claudesheets project."""
    from claudesheets.commands.init_cmd import run

    run(path)


@main.command('import')
@click.argument('xlsx', type=click.Path(exists=True, dir_okay=False))
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
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
)
def import_cmd(
    xlsx: str,
    archive: bool,
    flatten: bool,
    non_interactive: bool,
    project_path: str,
) -> None:
    """Read an .xlsx file into source form."""
    from claudesheets.commands.import_cmd import run

    run(
        xlsx_path=xlsx,
        project_path=project_path,
        archive=archive,
        flatten=flatten,
        non_interactive=non_interactive,
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
    help='Path to the claudesheets project.',
)
def build_cmd(out_path: str | None, project_path: str) -> None:
    """Compile source files into an .xlsx."""
    from claudesheets.commands.build_cmd import run

    run(project_path=project_path, out_path=out_path)


@main.command('recalc')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
)
@click.option(
    '--force',
    is_flag=True,
    help='Ignore the cache and re-run the calc engine.',
)
def recalc_cmd(project_path: str, force: bool) -> None:
    """Run the calc engine and cache calculated values."""
    from claudesheets.commands.recalc_cmd import run

    run(project_path=project_path, force=force)


@main.command('snapshot')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
)
@click.option(
    '--update',
    is_flag=True,
    help='Overwrite the saved snapshot with the current calculated values.',
)
def snapshot_cmd(project_path: str, update: bool) -> None:
    """Compare or update the golden-file snapshot of calculated values."""
    from claudesheets.commands.snapshot_cmd import run

    run(project_path=project_path, update=update)


@main.command('diff')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
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
    from claudesheets.commands.diff_cmd import run

    run(project_path=project_path, vs=vs)


@main.command('check')
@click.option(
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
)
def check_cmd(project_path: str) -> None:
    """Lint dangling refs, missing names, schema mismatches."""
    from claudesheets.commands.check_cmd import run

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
    help='Path to the claudesheets project.',
)
@click.argument('targets', nargs=-1, type=click.UNPROCESSED)
def test_cmd(project_path: str, targets: tuple[str, ...]) -> None:
    """Run the project's testsweet tests."""
    from claudesheets.commands.test_cmd import run

    run(project_path=project_path, targets=list(targets))


if __name__ == '__main__':  # pragma: no cover
    sys.exit(main())
