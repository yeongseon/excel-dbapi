from pathlib import Path
from typing import Any, cast

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.engines.base import TableData
from excel_dbapi.exceptions import ProgrammingError
from excel_dbapi.executor import SharedExecutor
from excel_dbapi.parser import parse_sql


def _create_workbook(
    path: Path,
    *,
    headers: list[object] | None = None,
    rows: list[list[object]] | None = None,
    sheet_name: str = "Sheet1",
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = sheet_name
    if headers is not None:
        sheet.append(headers)
    for row in rows or []:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


@pytest.mark.parametrize(
    "sql",
    [
        "CREATE TABLE (id INTEGER)",
        "CREATE TABLE t (,id INTEGER)",
        "CREATE TABLE t (id INTEGER,, name TEXT)",
    ],
)
def test_create_table_rejects_missing_name_and_empty_column_definitions(
    sql: str,
) -> None:
    with pytest.raises(ValueError):
        parse_sql(sql)


def test_create_table_rejects_single_trailing_comma() -> None:
    with pytest.raises(ValueError, match="empty column definition"):
        parse_sql("CREATE TABLE t (id INTEGER, name TEXT,)")

def test_execute_create_table_rejects_malformed_definitions(tmp_path: Path) -> None:
    file_path = tmp_path / "create-malformed.xlsx"
    _create_workbook(file_path, headers=["seed"], rows=[[1]])

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError, match="Table name is required"):
            cursor.execute("CREATE TABLE (id INTEGER)")
        with pytest.raises(
            ProgrammingError,
            match="Malformed column definitions: empty column definition found",
        ):
            cursor.execute("CREATE TABLE t (id INTEGER,, name TEXT)")


def test_execute_create_table_rejects_trailing_comma(tmp_path: Path) -> None:
    file_path = tmp_path / "create-trailing-comma.xlsx"
    _create_workbook(file_path, headers=["seed"], rows=[[1]])

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError, match="empty column definition"):
            cursor.execute("CREATE TABLE people (id INTEGER, name TEXT,)")

@pytest.mark.parametrize(
    "sql",
    [
        "CREATE TABLE __excel_meta__ (id INTEGER)",
        "DROP TABLE __excel_meta__",
        "ALTER TABLE __excel_meta__ ADD COLUMN x TEXT",
    ],
)
def test_reserved_metadata_sheet_rejects_ddl(tmp_path: Path, sql: str) -> None:
    file_path = tmp_path / "reserved-ddl.xlsx"
    _create_workbook(file_path, headers=["id"], rows=[[1]])

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        with pytest.raises(
            ProgrammingError,
            match="Cannot perform DDL on reserved metadata table '__excel_meta__'",
        ):
            cursor.execute(sql)


def test_fetch_methods_raise_before_execute(tmp_path: Path) -> None:
    file_path = tmp_path / "fetch-before-execute.xlsx"
    _create_workbook(file_path, headers=["id"], rows=[[1]])

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError, match="No result set"):
            cursor.fetchone()
        with pytest.raises(ProgrammingError, match="No result set"):
            cursor.fetchall()
        with pytest.raises(ProgrammingError, match="No result set"):
            cursor.fetchmany()


def test_fetch_methods_raise_after_update_and_ddl(tmp_path: Path) -> None:
    file_path = tmp_path / "fetch-after-write.xlsx"
    _create_workbook(file_path, headers=["id", "name"], rows=[[1, "Alice"]])

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()

        cursor.execute("UPDATE Sheet1 SET name = 'Bob' WHERE id = 1")
        with pytest.raises(ProgrammingError, match="No result set"):
            cursor.fetchall()

        cursor.execute("CREATE TABLE t (id INTEGER)")
        with pytest.raises(ProgrammingError, match="No result set"):
            cursor.fetchmany(1)


def test_alter_drop_column_is_case_insensitive(tmp_path: Path) -> None:
    file_path = tmp_path / "alter-drop-case-insensitive.xlsx"
    _create_workbook(
        file_path, headers=["id", "Name"], rows=[[1, "Alice"]], sheet_name="people"
    )

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("ALTER TABLE people DROP COLUMN name")
        cursor.execute("SELECT * FROM people")
        assert [str(col[0]) for col in (cursor.description or [])] == ["id"]
        assert cursor.fetchall() == [(1,)]


def test_alter_rename_column_is_case_insensitive(tmp_path: Path) -> None:
    file_path = tmp_path / "alter-rename-case-insensitive.xlsx"
    _create_workbook(
        file_path, headers=["id", "Name"], rows=[[1, "Alice"]], sheet_name="people"
    )

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("ALTER TABLE people RENAME COLUMN name TO full_name")
        cursor.execute("SELECT id, full_name FROM people")
        assert cursor.fetchall() == [(1, "Alice")]


class _NonTransactionalBackend:
    supports_transactions = False
    readonly = False

    def __init__(self) -> None:
        self._sheets = {
            "people": TableData(headers=["id", "name"], rows=[[1, "Alice"]])
        }

    def list_sheets(self) -> list[str]:
        return list(self._sheets)

    def read_sheet(self, name: str) -> TableData:
        return self._sheets[name]

    def write_sheet(self, name: str, data: TableData) -> None:
        self._sheets[name] = data


def test_non_transactional_metadata_read_failure_skips_lossy_rewrite(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import excel_dbapi.reflection as reflection

    backend = _NonTransactionalBackend()
    executor = SharedExecutor(cast(Any, backend), connection=object())
    write_calls = 0

    def _boom_read(*_: object, **__: object) -> list[dict[str, object]]:
        raise RuntimeError("metadata read failed")

    def _track_write(*_: object, **__: object) -> None:
        nonlocal write_calls
        write_calls += 1

    monkeypatch.setattr(reflection, "read_table_metadata", _boom_read)
    monkeypatch.setattr(reflection, "write_table_metadata", _track_write)

    with pytest.warns(UserWarning, match="skipping metadata update to avoid data loss"):
        executor.execute(parse_sql("ALTER TABLE people DROP COLUMN name"))

    assert write_calls == 0


def test_create_and_alter_share_type_normalization_and_validation() -> None:
    create_parsed = parse_sql("CREATE TABLE people (id INT, score FLOAT)")
    assert create_parsed["column_definitions"] == [
        {"name": "id", "type_name": "INTEGER"},
        {"name": "score", "type_name": "REAL"},
    ]

    alter_parsed = parse_sql("ALTER TABLE people ADD COLUMN age INT")
    assert alter_parsed["type_name"] == "INTEGER"

    with pytest.raises(ValueError, match="Unsupported CREATE TABLE column type"):
        parse_sql("CREATE TABLE bad (payload BLOB)")
    with pytest.raises(ValueError, match="Unsupported ALTER TABLE column type"):
        parse_sql("ALTER TABLE bad ADD COLUMN payload BLOB")
