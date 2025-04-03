import pytest
from unittest.mock import MagicMock, patch
from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import OperationalError


def test_connection_explicit(tmp_path):
    from openpyxl import Workbook

    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    wb.save(file_path)

    conn = ExcelConnection(file_path)
    conn.connect()
    assert conn._connected is True
    conn.close()
    assert conn._connected is False


def test_connection_double_close(tmp_path):
    from openpyxl import Workbook

    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    wb.save(file_path)

    conn = ExcelConnection(file_path)
    conn.connect()
    conn.close()
    # Multiple close should not raise error
    conn.close()
    assert conn._connected is False


def test_connection_commit_rollback(tmp_path):
    from openpyxl import Workbook

    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    wb.save(file_path)

    with ExcelConnection(file_path) as conn:
        conn.commit()  # Should do nothing
        conn.rollback()  # Should do nothing


@patch("excel_dbapi.connection.fetch_remote_file")
def test_connection_remote_file(mock_fetch):
    mock_fetch.return_value = b"fake-content"
    conn = ExcelConnection("https://example.com/test.xlsx")

    with pytest.raises(OperationalError):
        conn.connect()


def test_connection_with_custom_engine(tmp_path):
    from openpyxl import Workbook

    class DummyEngine:
        def load_workbook(self, file):
            self.loaded = True

        def close(self):
            self.closed = True

        def get_sheet(self, sheet_name):
            return None

        @property
        def workbook(self):
            return "dummy"

    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    wb.save(file_path)

    engine = DummyEngine()
    conn = ExcelConnection(file_path, engine=engine)
    conn.connect()
    assert engine.loaded is True
    conn.close()
    assert engine.closed is True
