import hashlib
import time
from pathlib import Path

from claudesheets.xlsx.writer import write_xlsx
from tests.fixtures.workbooks import write_simple_xlsx


def _sha256(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def test_two_builds_of_same_workbook_are_byte_identical(tmp_path: Path):
    from claudesheets.xlsx.reader import read_xlsx

    src = tmp_path / 'src.xlsx'
    write_simple_xlsx(src)
    wb = read_xlsx(src)

    a = tmp_path / 'a.xlsx'
    b = tmp_path / 'b.xlsx'
    write_xlsx(wb, a)
    time.sleep(1.1)  # ensure wall-clock would differ
    write_xlsx(wb, b)

    assert _sha256(a) == _sha256(b), (
        'xlsx writer is not deterministic: same Workbook produced '
        'different bytes 1s apart'
    )
