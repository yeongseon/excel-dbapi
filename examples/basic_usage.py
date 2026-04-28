# examples/basic_usage.py
# Demonstrates basic SELECT queries with excel-dbapi.

import tempfile
from pathlib import Path

import openpyxl

from excel_dbapi import connect


def _create_sample(path: Path) -> None:
    """Create a small sample workbook for this example."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name", "score"])
    ws.append([1, "Alice", 85])
    ws.append([2, "Bob", 72])
    ws.append([3, "Carol", 91])
    wb.save(str(path))


def main():
    with tempfile.TemporaryDirectory() as tmpdir:
        sample = Path(tmpdir) / "sample.xlsx"
        _create_sample(sample)

        with connect(str(sample)) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM Sheet1")
            for row in cursor.fetchall():
                print(row)


if __name__ == "__main__":
    main()
