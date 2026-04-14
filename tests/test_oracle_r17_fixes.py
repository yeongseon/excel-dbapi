from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import OperationalError


def _create_people_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet.append(["id", "Name", "phrase"])
    sheet.append([1, "Alice", "Stra\u00dfe"])
    sheet.append([2, "Bob", "Road"])
    workbook.save(path)
    workbook.close()


def test_commit_wraps_non_dbapi_backend_exceptions(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    file_path = tmp_path / "commit-wrap.xlsx"
    _create_people_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:

        def _boom_save() -> None:
            raise RuntimeError("save failed")

        monkeypatch.setattr(conn.engine, "save", _boom_save)
        with pytest.raises(OperationalError, match="save failed"):
            conn.commit()


def test_close_wraps_non_dbapi_backend_exceptions(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    file_path = tmp_path / "close-wrap.xlsx"
    _create_people_workbook(file_path)

    conn = ExcelConnection(str(file_path), engine="openpyxl", autocommit=True)

    def _boom_close() -> None:
        raise RuntimeError("close failed")

    monkeypatch.setattr(conn.engine, "close", _boom_close)
    with pytest.raises(OperationalError, match="close failed"):
        conn.close()


def test_select_update_insert_column_resolution_is_case_insensitive(
    tmp_path: Path,
) -> None:
    file_path = tmp_path / "column-casefold.xlsx"
    _create_people_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM Sheet1 ORDER BY id")
        assert cursor.fetchall() == [("Alice",), ("Bob",)]

        cursor.execute("UPDATE Sheet1 SET name = 'Carol' WHERE Name = 'Alice'")
        assert cursor.rowcount == 1
        cursor.execute("SELECT Name FROM Sheet1 WHERE id = 1")
        assert cursor.fetchall() == [("Carol",)]

        cursor.execute("INSERT INTO Sheet1 (id, name, phrase) VALUES (3, 'Dana', 'x')")
        cursor.execute("SELECT Name FROM Sheet1 WHERE id = 3")
        assert cursor.fetchall() == [("Dana",)]


def test_ilike_uses_unicode_casefold(tmp_path: Path) -> None:
    file_path = tmp_path / "unicode-ilike.xlsx"
    _create_people_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT phrase FROM Sheet1 WHERE phrase ILIKE '%STRASSE%'")
        assert cursor.fetchall() == [("Stra\u00dfe",)]
