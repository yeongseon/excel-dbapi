import pytest

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import OperationalError


def test_connection_local(tmp_path):
    from openpyxl import Workbook

    file_path = tmp_path / "test.xlsx"
    wb = Workbook()
    wb.save(file_path)

    with ExcelConnection(file_path) as conn:
        assert conn.engine.workbook is not None

    with pytest.raises(OperationalError):
        _ = conn.engine.workbook


def test_connection_invalid_file():
    with pytest.raises(OperationalError):
        with ExcelConnection("nonexistent.xlsx"):
            pass
