import tempfile
from contextlib import contextmanager
from pathlib import Path

import openpyxl
from openpyxl.worksheet.table import Table, TableColumn
from testsweet import params, test

from claudesheets.xlsx.reader import read_xlsx
from claudesheets.xlsx.writer import write_xlsx


@contextmanager
def _tmp_path():
    with tempfile.TemporaryDirectory() as td:
        yield Path(td)


def _wb_with_table(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'S'
    ws['A1'], ws['B1'], ws['C1'] = 'Region', 'Q1', 'Q2'
    for i, region in enumerate(['North', 'South', 'East'], start=2):
        ws.cell(row=i, column=1, value=region)
        ws.cell(row=i, column=2, value=10 * i)
        ws.cell(row=i, column=3, value=20 * i)

    cols = [
        TableColumn(id=1, name='Region'),
        TableColumn(id=2, name='Q1'),
        TableColumn(id=3, name='Q2'),
    ]
    t = Table(
        displayName='Sales',
        name='Sales',
        ref='A1:C4',
        headerRowCount=1,
        totalsRowCount=0,
        tableColumns=cols,
    )
    ws.add_table(t)
    wb.save(path)


@test
def reader_picks_up_table():
    with _tmp_path() as tmp_path:
        src = tmp_path / 'in.xlsx'
        _wb_with_table(src)
        wb = read_xlsx(src)
        s = wb.sheet('S')
        assert len(s.tables) == 1
        assert s.tables[0].name == 'Sales'
        assert s.tables[0].ref == 'A1:C4'
        assert [c.name for c in s.tables[0].columns] == [
            'Region',
            'Q1',
            'Q2',
        ]


@test
def table_round_trips_through_writer():
    with _tmp_path() as tmp_path:
        src = tmp_path / 'in.xlsx'
        _wb_with_table(src)
        out = tmp_path / 'out.xlsx'
        write_xlsx(read_xlsx(src), out)
        s = read_xlsx(out).sheet('S')
        assert len(s.tables) == 1
        assert s.tables[0].name == 'Sales'
        assert [c.name for c in s.tables[0].columns] == [
            'Region',
            'Q1',
            'Q2',
        ]


@test
@params([(h, t) for h in (0, 1, 2) for t in (0, 1)])
def table_round_trips_with_varying_header_and_totals_rows(
    header_count: int, totals_count: int
):
    """Cover the {0,1,2} header / {0,1} totals matrix.

    Excel allows 0 (rare, for borderless data tables), 1 (default),
    or 2 (header + sub-header) header rows. Totals row is 0 or 1.
    """
    with _tmp_path() as tmp_path:
        src = tmp_path / f'in_h{header_count}_t{totals_count}.xlsx'
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'S'
        # Layout: enough rows for headers + 3 data rows + optional totals.
        n_rows = max(header_count, 1) + 3 + totals_count
        for r in range(1, n_rows + 1):
            ws.cell(row=r, column=1, value=f'r{r}c1')
            ws.cell(row=r, column=2, value=r * 10)
        ref = f'A1:B{n_rows}'

        cols = [
            TableColumn(id=1, name='Label'),
            TableColumn(
                id=2,
                name='Value',
                totalsRowFunction='sum' if totals_count else None,
            ),
        ]
        t = Table(
            displayName='T',
            name='T',
            ref=ref,
            headerRowCount=header_count,
            totalsRowCount=totals_count,
            tableColumns=cols,
        )
        ws.add_table(t)
        wb.save(src)

        out = tmp_path / f'out_h{header_count}_t{totals_count}.xlsx'
        write_xlsx(read_xlsx(src), out)
        s = read_xlsx(out).sheet('S')

        assert len(s.tables) == 1
        assert s.tables[0].header_row_count == header_count
        assert s.tables[0].totals_row_count == totals_count
