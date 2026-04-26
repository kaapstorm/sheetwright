"""claudesheets CLI entry point."""

from __future__ import annotations

import sys

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
    '--project',
    'project_path',
    type=click.Path(file_okay=False),
    default='.',
    help='Path to the claudesheets project.',
)
def import_cmd(xlsx: str, archive: bool, project_path: str) -> None:
    """Read an .xlsx file into source form."""
    from claudesheets.commands.import_cmd import run

    run(xlsx_path=xlsx, project_path=project_path, archive=archive)


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


if __name__ == '__main__':  # pragma: no cover
    sys.exit(main())
