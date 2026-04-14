from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import NotSupportedError, ProgrammingError


def _create_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.append(["id", "name"])
    ws.append([1, "Alice"])
    wb.save(path)


def _create_headerless_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    wb.save(path)


def test_unknown_sheet_raises_programming_error(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path)) as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError):
            cursor.execute("SELECT * FROM MissingSheet")


def test_unknown_column_raises_programming_error(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path)) as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError):
            cursor.execute("SELECT missing FROM Sheet1")


def test_unknown_order_by_column_raises_programming_error(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path)) as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError):
            cursor.execute("SELECT id FROM Sheet1 ORDER BY missing")


def test_unsupported_operator_raises_not_supported(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path)) as conn:
        cursor = conn.cursor()
        with pytest.raises(NotSupportedError):
            cursor.execute("SELECT id FROM Sheet1 WHERE id LIKE 1")


def test_openpyxl_insert_mismatch_columns(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError):
            cursor.execute("INSERT INTO Sheet1 (id) VALUES (1, 'A')")


def test_openpyxl_create_drop_errors(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE NewSheet (id)")
        with pytest.raises(ProgrammingError):
            cursor.execute("CREATE TABLE NewSheet (id)")
        cursor.execute("DROP TABLE NewSheet")
        with pytest.raises(ProgrammingError):
            cursor.execute("DROP TABLE NewSheet")


def test_pandas_create_drop_errors(tmp_path: Path) -> None:
    file_path = tmp_path / "sample.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="pandas") as conn:
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE NewSheet (id)")
        with pytest.raises(ProgrammingError):
            cursor.execute("CREATE TABLE NewSheet (id)")
        cursor.execute("DROP TABLE NewSheet")
        with pytest.raises(ProgrammingError):
            cursor.execute("DROP TABLE NewSheet")


def test_select_column_on_headerless_sheet_raises_programming_error(
    tmp_path: Path,
) -> None:
    file_path = tmp_path / "headerless-select.xlsx"
    _create_headerless_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(
            ProgrammingError, match="No columns defined in sheet 'Sheet1'"
        ):
            cursor.execute("SELECT id FROM Sheet1")


def test_select_star_on_headerless_sheet_returns_empty_rows(tmp_path: Path) -> None:
    file_path = tmp_path / "headerless-select-star.xlsx"
    _create_headerless_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        assert cursor.fetchall() == []


def test_update_headerless_sheet_raises_programming_error(tmp_path: Path) -> None:
    file_path = tmp_path / "headerless-update.xlsx"
    _create_headerless_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(
            ProgrammingError, match="No columns defined in sheet 'Sheet1'"
        ):
            cursor.execute("UPDATE Sheet1 SET id = 1")


def test_delete_where_on_headerless_sheet_raises_programming_error(
    tmp_path: Path,
) -> None:
    file_path = tmp_path / "headerless-delete.xlsx"
    _create_headerless_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(
            ProgrammingError, match="No columns defined in sheet 'Sheet1'"
        ):
            cursor.execute("DELETE FROM Sheet1 WHERE id = 1")


def test_drop_only_remaining_sheet_raises_programming_error(tmp_path: Path) -> None:
    file_path = tmp_path / "drop-last-sheet.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(
            ProgrammingError, match="Cannot drop the only remaining sheet"
        ):
            cursor.execute("DROP TABLE Sheet1")
