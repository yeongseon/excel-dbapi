from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection


def _create_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "score"])
    ws.append([1, 5])
    ws.append([2, 10])
    ws.append([3, 15])
    wb.save(path)


def test_openpyxl_comparison_operators(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM Sheet1 WHERE score > 5")
        assert cursor.fetchall() == [(2,), (3,)]
        cursor.execute("SELECT id FROM Sheet1 WHERE score <= 10")
        assert cursor.fetchall() == [(1,), (2,)]


def test_pandas_comparison_operators(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    df = pd.DataFrame(
        [{"id": 1, "score": 5}, {"id": 2, "score": 10}, {"id": 3, "score": 15}]
    )
    df.to_excel(file_path, index=False, sheet_name="Sheet1")

    with ExcelConnection(str(file_path), engine="pandas") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM Sheet1 WHERE score > 5")
        assert cursor.fetchall() == [(2,), (3,)]
        cursor.execute("SELECT id FROM Sheet1 WHERE score <= 10")
        assert cursor.fetchall() == [(1,), (2,)]
