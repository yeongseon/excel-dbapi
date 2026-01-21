from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from excel_dbapi.connection import ExcelConnection


def _create_sample_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    wb.save(path)


def _create_multi_row_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    ws.append([2, "Bob"])
    ws.append([3, "Cara"])
    wb.save(path)


def test_openpyxl_insert_and_executemany(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_sample_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (2, 'Bob')")
        assert cursor.rowcount == 1

        cursor.executemany(
            "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
            [(3, "Cora"), (4, "Dane")],
        )
        assert cursor.rowcount == 2

    wb = load_workbook(file_path, data_only=True)
    rows = list(wb["Sheet1"].iter_rows(values_only=True))
    assert rows[-3:] == [(2, "Bob"), (3, "Cora"), (4, "Dane")]


def test_openpyxl_create_and_drop_table(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_sample_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE NewSheet (col1, col2)")
        cursor.execute("DROP TABLE NewSheet")

    wb = load_workbook(file_path, data_only=True)
    assert "NewSheet" not in wb.sheetnames


def test_pandas_insert_and_create(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    df = pd.DataFrame([{"id": 1, "name": "Alice"}])
    df.to_excel(file_path, index=False, sheet_name="Sheet1")

    with ExcelConnection(str(file_path), engine="pandas", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (2, 'Bob')")
        cursor.execute("CREATE TABLE Extra (col1, col2)")

    data = pd.read_excel(file_path, sheet_name=None)
    assert len(data["Sheet1"]) == 2
    assert set(data["Extra"].columns) == {"col1", "col2"}


def test_openpyxl_update_delete_and_rollback(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_sample_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=False) as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
        assert cursor.rowcount == 1
        cursor.execute("DELETE FROM Sheet1 WHERE id = 1")
        assert cursor.rowcount == 1
        conn.rollback()

    wb = load_workbook(file_path, data_only=True)
    rows = list(wb["Sheet1"].iter_rows(values_only=True))
    assert rows[1] == (1, "Alice")


def test_pandas_update_and_delete(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    df = pd.DataFrame([{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}])
    df.to_excel(file_path, index=False, sheet_name="Sheet1")

    with ExcelConnection(str(file_path), engine="pandas", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 2")
        assert cursor.rowcount == 1
        cursor.execute("DELETE FROM Sheet1 WHERE id = 1")
        assert cursor.rowcount == 1

    data = pd.read_excel(file_path, sheet_name=None)
    assert list(data["Sheet1"]["name"]) == ["Ann"]


def test_select_order_limit_with_where_openpyxl(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_multi_row_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM Sheet1 WHERE id >= 2 ORDER BY id DESC LIMIT 1")
        results = cursor.fetchall()
        assert results == [(3, "Cara")]


def test_select_order_limit_with_where_pandas(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    df = pd.DataFrame(
        [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}]
    )
    df.to_excel(file_path, index=False, sheet_name="Sheet1")

    with ExcelConnection(str(file_path), engine="pandas", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM Sheet1 WHERE id >= 2 ORDER BY id DESC LIMIT 1")
        results = cursor.fetchall()
        assert results == [(3, "Cara")]
