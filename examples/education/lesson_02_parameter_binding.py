# examples/education/lesson_02_parameter_binding.py
# Lesson 2: Use parameter binding for safe INSERT queries.

import tempfile
from pathlib import Path

import openpyxl

from excel_dbapi import connect


def _create_sample(path: Path) -> None:
    """Create a small sample workbook for this lesson."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    ws.append([2, "Bob"])
    wb.save(str(path))


with tempfile.TemporaryDirectory() as tmpdir:
    sample = Path(tmpdir) / "sample.xlsx"
    _create_sample(sample)

    with connect(str(sample), autocommit=False) as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)", (101, "Student"))
        conn.commit()
