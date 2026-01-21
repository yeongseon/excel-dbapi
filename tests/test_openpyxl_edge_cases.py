from pathlib import Path

from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection


def test_select_on_empty_sheet_returns_empty(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "EmptySheet"
    wb.save(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM EmptySheet")
        assert cursor.fetchall() == []
