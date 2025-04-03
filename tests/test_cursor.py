import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import InterfaceError


def test_cursor_fetch_methods():
    conn = ExcelConnection("tests/data/sample.xlsx")
    cursor = conn.cursor()
    cursor._results = [{"id": 1}, {"id": 2}]
    assert cursor.fetchone() == {"id": 1}
    assert cursor.fetchall() == [{"id": 2}]
    cursor.close()
    conn.close()


def test_cursor_closed():
    conn = ExcelConnection("tests/data/sample.xlsx")
    cursor = conn.cursor()
    cursor.close()
    with pytest.raises(InterfaceError):
        cursor.execute("SELECT * FROM [Sheet1$]")
    conn.close()
