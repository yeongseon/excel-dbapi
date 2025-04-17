import pytest
from excel_dbapi.connection import ExcelConnection


def test_connection_open_and_close():
    conn = ExcelConnection("tests/data/sample.xlsx")
    assert conn.closed is False
    conn.close()
    assert conn.closed is True


def test_connection_cursor():
    conn = ExcelConnection("tests/data/sample.xlsx")
    cursor = conn.cursor()
    assert cursor is not None
    conn.close()
    with pytest.raises(Exception):
        conn.cursor()
