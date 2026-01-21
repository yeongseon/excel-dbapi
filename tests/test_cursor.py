import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import NotSupportedError, ProgrammingError


def test_cursor_execute_and_fetchall():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        results = cursor.fetchall()
        assert isinstance(results, list)
        assert isinstance(results[0], tuple)
        assert cursor.description is not None
        assert cursor.rowcount == len(results)


def test_cursor_fetchone():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        row = cursor.fetchone()
        assert isinstance(row, tuple)

def test_cursor_closed():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.close()
        with pytest.raises(Exception):
            cursor.execute("SELECT * FROM Sheet1")


def test_cursor_error_translation():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError):
            cursor.execute("INVALID SQL")

        with pytest.raises(NotSupportedError):
            cursor.execute("SELECT * FROM Sheet1 WHERE id LIKE 1")
