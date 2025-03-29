import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import ProgrammingError


def test_cursor_execute_fetchall(tmp_path):
    from openpyxl import Workbook

    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Age"])
    ws.append(["Alice", 25])
    ws.append(["Bob", 30])
    wb.save(file_path)

    with ExcelConnection(file_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet")
        result = cursor.fetchall()
        assert result == [{"Name": "Alice", "Age": 25}, {"Name": "Bob", "Age": 30}]
        assert cursor.rowcount == 2


def test_cursor_fetchone(tmp_path):
    from openpyxl import Workbook

    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Age"])
    ws.append(["Alice", 25])
    wb.save(file_path)

    with ExcelConnection(file_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet")
        row = cursor.fetchone()
        assert row == {"Name": "Alice", "Age": 25}
        assert cursor.fetchone() is None


def test_cursor_invalid_query(tmp_path):
    from openpyxl import Workbook

    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    wb.save(file_path)

    with ExcelConnection(file_path) as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError):
            cursor.execute("INSERT INTO Sheet")
