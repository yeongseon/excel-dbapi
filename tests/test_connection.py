import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import InterfaceError, NotSupportedError


def test_connection_with_pandas_engine():
    conn = ExcelConnection("tests/data/sample.xlsx", engine="pandas")
    assert conn.engine.__class__.__name__ == "PandasEngine"
    conn.close()
    assert conn.closed is True


def test_connection_with_openpyxl_engine():
    conn = ExcelConnection("tests/data/sample.xlsx", engine="openpyxl")
    assert conn.engine.__class__.__name__ == "OpenpyxlEngine"
    conn.close()
    assert conn.closed is True


def test_connection_with_invalid_engine():
    with pytest.raises(ValueError) as e:
        ExcelConnection("tests/data/sample.xlsx", engine="invalid")
    assert "Unsupported engine" in str(e.value)


def test_connection_context_manager():
    with ExcelConnection("tests/data/sample.xlsx", engine="pandas") as conn:
        assert conn.closed is False
    assert conn.closed is True


def test_connection_closed_error():
    conn = ExcelConnection("tests/data/sample.xlsx", engine="pandas")
    conn.close()
    with pytest.raises(InterfaceError):
        conn.cursor()


def test_connection_commit_rollback_not_supported():
    conn = ExcelConnection("tests/data/sample.xlsx", engine="openpyxl")
    with pytest.raises(NotSupportedError):
        conn.commit()
    with pytest.raises(NotSupportedError):
        conn.rollback()
    conn.close()


def test_engine_name_pandas():
    conn = ExcelConnection("tests/data/sample.xlsx", engine="pandas")
    assert conn.engine_name == "PandasEngine"
    conn.close()


def test_engine_name_openpyxl():
    conn = ExcelConnection("tests/data/sample.xlsx", engine="openpyxl")
    assert conn.engine_name == "OpenpyxlEngine"
    conn.close()


def test_connection_str_and_repr():
    conn = ExcelConnection("tests/data/sample.xlsx", engine="pandas")
    conn_str = str(conn)
    conn_repr = repr(conn)

    assert "ExcelConnection" in conn_str
    assert "pandas" in conn_str or "PandasEngine" in conn_str
    assert conn_str == conn_repr

    conn.close()
    assert "closed=True" in str(conn)
