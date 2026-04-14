from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import DataError, NotSupportedError, ProgrammingError
from excel_dbapi.reflection import METADATA_SHEET, read_table_metadata, write_table_metadata


def _create_workbook(path: Path, headers: list[object], rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def test_pandas_backend_rejects_blank_headers(tmp_path: Path) -> None:
    file_path = tmp_path / "pandas-unnamed.xlsx"
    _create_workbook(file_path, ["id", None], [[1, "Alice"]])

    with pytest.raises(DataError, match="Empty or None header"):
        ExcelConnection(str(file_path), engine="pandas")


def test_pandas_backend_rejects_duplicate_headers(tmp_path: Path) -> None:
    file_path = tmp_path / "pandas-duplicate.xlsx"
    _create_workbook(file_path, ["id", "id"], [[1, 2]])

    with pytest.raises(DataError, match="Duplicate header"):
        ExcelConnection(str(file_path), engine="pandas")


def test_pandas_backend_rejects_data_only_false(tmp_path: Path) -> None:
    file_path = tmp_path / "pandas-data-only.xlsx"
    _create_workbook(file_path, ["id"], [[1]])

    with pytest.raises(NotSupportedError, match="does not support data_only=False"):
        ExcelConnection(str(file_path), engine="pandas", data_only=False)


def test_pandas_lastrowid_matches_openpyxl(tmp_path: Path) -> None:
    openpyxl_file = tmp_path / "openpyxl-lastrowid.xlsx"
    pandas_file = tmp_path / "pandas-lastrowid.xlsx"
    _create_workbook(openpyxl_file, ["id", "name"], [[1, "Alice"]])
    _create_workbook(pandas_file, ["id", "name"], [[1, "Alice"]])

    with ExcelConnection(str(openpyxl_file), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Sheet1 VALUES (2, 'Bob')")
        openpyxl_lastrowid = cursor.lastrowid

    with ExcelConnection(str(pandas_file), engine="pandas") as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Sheet1 VALUES (2, 'Bob')")
        pandas_lastrowid = cursor.lastrowid

    assert openpyxl_lastrowid == 3
    assert pandas_lastrowid == openpyxl_lastrowid


def test_ddl_updates_reflection_metadata(tmp_path: Path) -> None:
    file_path = tmp_path / "ddl-metadata.xlsx"
    _create_workbook(file_path, ["seed"], [[1]])

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()

        cursor.execute("CREATE TABLE people (id, name)")
        metadata = read_table_metadata(conn, "people")
        assert metadata is not None
        assert [entry["name"] for entry in metadata] == ["id", "name"]

        cursor.execute("ALTER TABLE people ADD COLUMN email TEXT")
        metadata = read_table_metadata(conn, "people")
        assert metadata is not None
        assert [entry["name"] for entry in metadata] == ["id", "name", "email"]

        cursor.execute("ALTER TABLE people RENAME COLUMN name TO full_name")
        metadata = read_table_metadata(conn, "people")
        assert metadata is not None
        assert [entry["name"] for entry in metadata] == ["id", "full_name", "email"]

        cursor.execute("ALTER TABLE people DROP COLUMN email")
        metadata = read_table_metadata(conn, "people")
        assert metadata is not None
        assert [entry["name"] for entry in metadata] == ["id", "full_name"]

        cursor.execute("DROP TABLE people")
        assert read_table_metadata(conn, "people") is None


def test_drop_table_guard_ignores_metadata_sheet(tmp_path: Path) -> None:
    file_path = tmp_path / "drop-last-user-sheet.xlsx"
    _create_workbook(file_path, ["id"], [[1]])

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        write_table_metadata(
            conn,
            "Sheet1",
            [
                {
                    "name": "id",
                    "type_name": "INTEGER",
                    "nullable": False,
                    "primary_key": True,
                }
            ],
        )
        assert METADATA_SHEET in conn.engine.list_sheets()

        cursor = conn.cursor()
        with pytest.raises(ProgrammingError, match="Cannot drop the only remaining sheet"):
            cursor.execute("DROP TABLE Sheet1")
