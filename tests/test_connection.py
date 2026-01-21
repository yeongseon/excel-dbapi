import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import NotSupportedError


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


def test_rollback_autocommit_raises():
    with ExcelConnection("tests/data/sample.xlsx", autocommit=True) as conn:
        with pytest.raises(NotSupportedError):
            conn.rollback()
