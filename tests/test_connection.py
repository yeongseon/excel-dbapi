import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import DatabaseError, NotSupportedError, ProgrammingError


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


@pytest.mark.parametrize(
    ("raised", "expected"),
    [
        (ValueError("bad query"), ProgrammingError),
        (NotImplementedError("not supported"), NotSupportedError),
        (RuntimeError("boom"), DatabaseError),
    ],
)
def test_connection_execute_maps_exceptions(raised, expected):
    conn = ExcelConnection("tests/data/sample.xlsx")

    def _raise(*args, **kwargs):
        del args, kwargs
        raise raised

    conn._executor.execute_with_params = _raise

    with pytest.raises(expected):
        conn.execute("SELECT * FROM Sheet1")

    conn.close()
