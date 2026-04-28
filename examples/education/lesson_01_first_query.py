# examples/education/lesson_01_first_query.py
# Lesson 1: Run your first SELECT query against an Excel file.

import tempfile
from pathlib import Path

import openpyxl

from excel_dbapi import connect


def _create_sample(path: Path) -> None:
    """Create a small sample workbook for this lesson."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name", "score"])
    ws.append([1, "Alice", 85])
    ws.append([2, "Bob", 72])
    ws.append([3, "Carol", 91])
    ws.append([4, "Dave", 68])
    ws.append([5, "Eve", 95])
    wb.save(str(path))


with tempfile.TemporaryDirectory() as tmpdir:
    sample = Path(tmpdir) / "sample.xlsx"
    _create_sample(sample)

    with connect(str(sample)) as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, name FROM Sheet1 ORDER BY id LIMIT 5")
        for row in cur.fetchall():
            print(row)
