# examples/write_operations.py
# Demonstrates INSERT, UPDATE, DELETE, and executemany.

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

        with connect(str(sample)) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (10, 'Zoe')")
            cursor.execute("UPDATE Sheet1 SET name = 'Zoey' WHERE id = 10")
            cursor.execute("DELETE FROM Sheet1 WHERE id = 10")

            cursor.executemany(
                "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
                [(11, "Mina"), (12, "Noah")],
            )


if __name__ == "__main__":
    main()
