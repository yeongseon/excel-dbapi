import pytest
from excel_dbapi.connection import ExcelConnection


def test_cursor_execute_and_fetchall():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        results = cursor.fetchall()
        assert isinstance(results, list)
        assert isinstance(results[0], dict)


def test_cursor_fetchone():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        row = cursor.fetchone()
        assert isinstance(row, dict)

def test_cursor_closed():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.close()
        with pytest.raises(Exception):
            cursor.execute("SELECT * FROM Sheet1")
