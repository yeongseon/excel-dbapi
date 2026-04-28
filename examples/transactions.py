# examples/transactions.py
# Demonstrates rollback with autocommit=False.

import tempfile
from pathlib import Path

import openpyxl

from excel_dbapi import connect


def _create_sample(path: Path) -> None:
    """Create a small sample workbook for this example."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    ws.append([2, "Bob"])
    wb.save(str(path))


def main():
    with tempfile.TemporaryDirectory() as tmpdir:
        sample = Path(tmpdir) / "sample.xlsx"
        _create_sample(sample)

        with connect(str(sample), autocommit=False) as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
            conn.rollback()


if __name__ == "__main__":
    main()
