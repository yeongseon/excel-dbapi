from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import ProgrammingError


def _create_workbook(
    path: Path,
    *,
    headers: list[object],
    rows: list[list[object]],
    sheet_name: str,
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = sheet_name
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def test_fetchone_raises_after_connection_close(tmp_path: Path) -> None:
    file_path = tmp_path / "closed-connection-fetch.xlsx"
    _create_workbook(
        file_path,
        headers=["id", "name"],
        rows=[[1, "Alice"]],
        sheet_name="people",
    )

    conn = ExcelConnection(str(file_path), engine="openpyxl")
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM people")
    conn.close()

    with pytest.raises(ProgrammingError, match="Cannot operate on a closed connection"):
        cursor.fetchone()


def test_failed_execute_clears_prior_result_set(tmp_path: Path) -> None:
    file_path = tmp_path / "stale-results.xlsx"
    _create_workbook(
        file_path,
        headers=["id", "name"],
        rows=[[1, "Alice"], [2, "Bob"]],
        sheet_name="people",
    )

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM people ORDER BY id")
        assert cursor.fetchone() == (1,)

        with pytest.raises(ProgrammingError):
            cursor.execute("INVALID SQL")

        with pytest.raises(ProgrammingError, match="No result set"):
            cursor.fetchone()


def test_ascii_identifiers_work_end_to_end(tmp_path: Path) -> None:
    file_path = tmp_path / "ascii-identifiers.xlsx"
    _create_workbook(
        file_path,
        headers=["user_id", "full_name"],
        rows=[[1, "Alice"], [2, "Bob"]],
        sheet_name="users",
    )

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT user_id, full_name FROM users ORDER BY user_id")
        assert cursor.fetchall() == [(1, "Alice"), (2, "Bob")]


@pytest.mark.xfail(
    reason=(
        "Quoted identifiers are not yet supported for table/column resolution; "
        "double quotes are currently parsed as string literals"
    ),
    strict=False,
)
def test_spaced_identifiers_quoted_are_currently_not_supported(tmp_path: Path) -> None:
    file_path = tmp_path / "spaced-identifiers.xlsx"
    _create_workbook(
        file_path,
        headers=["id", "full name"],
        rows=[[1, "Alice"]],
        sheet_name="People Sheet",
    )

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT "full name" FROM "People Sheet"')
        assert cursor.fetchall() == [("Alice",)]
