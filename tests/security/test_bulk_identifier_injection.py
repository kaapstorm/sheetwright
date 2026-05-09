import csv
import io
import sqlite3
import tempfile
from contextlib import contextmanager
from pathlib import Path
from unittest import mock

from testsweet import catch_exceptions, test

from sheetwright.bulk import build_bulk_cache
from sheetwright.exceptions import BulkInvalidIdentifierError


@contextmanager
def _project():
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td)
        proj = tmp_path / 'p'
        (proj / 'data').mkdir(parents=True)
        (proj / '.sheetwright').mkdir(parents=True)
        yield proj


@test
def injection_payload_in_header_raises():
    # The header 'id, name"); DROP TABLE foo; --' contains a quote and
    # semicolons, both excluded by _IDENT_RE.
    with _project() as p:
        (p / 'data' / 'series.csv').write_text(
            'id, name"); DROP TABLE foo; --\n1,evil\n'
        )
        with catch_exceptions() as excs:
            build_bulk_cache(p)
    assert excs and isinstance(excs[0], BulkInvalidIdentifierError)


@test
def backtick_in_header_raises():
    with _project() as p:
        (p / 'data' / 'series.csv').write_text('`id`\n1\n')
        with catch_exceptions() as excs:
            build_bulk_cache(p)
    assert excs and isinstance(excs[0], BulkInvalidIdentifierError)


@test
def double_quote_in_header_raises():
    # Write via csv.writer so that the double-quote is embedded inside
    # the field value (csv.writer encodes 'col"name' as '"col""name"').
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(['col"name'])
    w.writerow(['1'])
    with _project() as p:
        (p / 'data' / 'series.csv').write_text(buf.getvalue())
        with catch_exceptions() as excs:
            build_bulk_cache(p)
    assert excs and isinstance(excs[0], BulkInvalidIdentifierError)


@test
def header_starting_with_digit_raises():
    with _project() as p:
        (p / 'data' / 'series.csv').write_text('1id\n1\n')
        with catch_exceptions() as excs:
            build_bulk_cache(p)
    assert excs and isinstance(excs[0], BulkInvalidIdentifierError)


@test
def header_that_is_single_digit_raises():
    with _project() as p:
        (p / 'data' / 'series.csv').write_text('1\n1\n')
        with catch_exceptions() as excs:
            build_bulk_cache(p)
    assert excs and isinstance(excs[0], BulkInvalidIdentifierError)


@test
def header_longer_than_63_chars_raises():
    long_header = 'a' * 64
    with _project() as p:
        (p / 'data' / 'series.csv').write_text(f'{long_header}\n1\n')
        with catch_exceptions() as excs:
            build_bulk_cache(p)
    assert excs and isinstance(excs[0], BulkInvalidIdentifierError)


@test
def table_name_validation_catches_bad_sanitised_output():
    # table_name_for already sanitises filenames, but _validate_identifier
    # is applied to its output as belt-and-braces. We monkey-patch
    # table_name_for to return '1bad' (starts with digit) to exercise the
    # path where a future change to table_name_for could produce an invalid
    # identifier that slips through without validation.
    with _project() as p:
        (p / 'data' / 'series.csv').write_text('id\n1\n')
        with mock.patch(
            'sheetwright.bulk.table_name_for', return_value='1bad'
        ):
            with catch_exceptions() as excs:
                build_bulk_cache(p)
    assert excs and isinstance(excs[0], BulkInvalidIdentifierError)


@test
def normal_csv_builds_cache_successfully():
    with _project() as p:
        (p / 'data' / 'series.csv').write_text(
            'id,name,value\n1,alpha,10\n2,beta,20\n'
        )
        db_path = build_bulk_cache(p)
        assert db_path.exists()

        conn = sqlite3.connect(db_path)
        rows = conn.execute(
            'SELECT id, name, value FROM series ORDER BY id'
        ).fetchall()
        conn.close()
        assert rows == [('1', 'alpha', '10'), ('2', 'beta', '20')]
