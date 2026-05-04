from click.testing import CliRunner
from testsweet import test

from sheetwright.cli import main


@test
def help_lists_subcommands():
    runner = CliRunner()
    result = runner.invoke(main, ['--help'])
    assert result.exit_code == 0
    for cmd in ('init', 'import', 'build'):
        assert cmd in result.output


@test
def unknown_subcommand_errors():
    runner = CliRunner()
    result = runner.invoke(main, ['bogus'])
    assert result.exit_code != 0
