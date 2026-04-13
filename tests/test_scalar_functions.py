from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import ProgrammingError


def _create_scalar_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "t"
    sheet.append(
        [
            "id",
            "name",
            "city",
            "nickname",
            "num_text",
            "float_text",
            "date_value",
            "date_text",
            "first",
            "last",
        ]
    )
    sheet.append(
        [
            1,
            "  alice  ",
            "new york",
            None,
            "42",
            "3.5",
            date(2024, 1, 15),
            "2024-01-15",
            "Ada",
            "Lovelace",
        ]
    )
    sheet.append(
        [
            2,
            "bob",
            "london",
            "bob",
            "7",
            "2.0",
            datetime(2023, 12, 5, 9, 30, 0),
            "2023-12-05T09:30:00",
            "Bob",
            "Builder",
        ]
    )
    sheet.append(
        [
            3,
            None,
            "paris",
            None,
            "abc",
            "bad",
            None,
            None,
            None,
            "Solo",
        ]
    )

    teams = workbook.create_sheet("teams")
    teams.append(["id", "team"])
    teams.append([1, "ENG"])
    teams.append([2, "OPS"])
    teams.append([3, "SALES"])

    workbook.save(path)


def _query(path: Path, sql: str) -> tuple[list[tuple[object, ...]], list[str]]:
    with ExcelConnection(str(path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(sql)
        rows = cursor.fetchall()
        description = [str(item[0]) for item in cursor.description or []]
        return rows, description


def test_scalar_functions_in_select(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_select.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT UPPER(name), LOWER(city), TRIM(name), LENGTH(name), "
        "SUBSTR('hello', 2, 3), SUBSTRING('world', 2, 3), CONCAT(first, ' ', last) "
        "FROM t WHERE id = 1",
    )
    assert rows == [("  ALICE  ", "new york", "alice", 9, "ell", "orl", "Ada Lovelace")]


def test_coalesce_and_nullif(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_coalesce_nullif.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT COALESCE(nickname, TRIM(name), 'fallback'), NULLIF(city, 'london') "
        "FROM t ORDER BY id",
    )
    assert rows == [
        ("alice", "new york"),
        ("bob", None),
        ("fallback", "paris"),
    ]


def test_scalar_functions_in_where(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_where.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT id FROM t WHERE UPPER(TRIM(name)) = 'ALICE' ORDER BY id",
    )
    assert rows == [(1,)]


def test_scalar_functions_in_order_by(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_order_by.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT id FROM t ORDER BY LOWER(city) DESC",
    )
    assert rows == [(3,), (1,), (2,)]


def test_nested_scalar_functions(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_nested.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT UPPER(TRIM(name)) FROM t WHERE id = 1",
    )
    assert rows == [("ALICE",)]


def test_cast_to_core_types(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_cast_types.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT CAST(num_text AS INTEGER), CAST(float_text AS REAL), "
        "CAST(date_value AS TEXT), CAST(date_text AS DATE) "
        "FROM t WHERE id = 1",
    )

    assert rows[0][0] == 42
    assert rows[0][1] == 3.5
    assert str(rows[0][2]).startswith("2024-01-15")
    assert rows[0][3] == date(2024, 1, 15)


def test_cast_null_is_null(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_cast_null.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT CAST(NULL AS INTEGER), CAST(NULL AS REAL), CAST(NULL AS TEXT), CAST(NULL AS DATE) "
        "FROM t WHERE id = 1",
    )
    assert rows == [(None, None, None, None)]


def test_cast_invalid_value_raises_programming_error(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_cast_invalid.xlsx"
    _create_scalar_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError, match="Cannot cast value"):
            cursor.execute("SELECT CAST(num_text AS INTEGER) FROM t WHERE id = 3")

        with pytest.raises(ProgrammingError, match="Cannot cast value"):
            cursor.execute("SELECT CAST('01/15/2024' AS DATE) FROM t WHERE id = 1")


def test_concatenation_operator_pipe_pipe(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_concat_operator.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT first || ' ' || last FROM t WHERE id = 1",
    )
    assert rows == [("Ada Lovelace",)]

    rows, _ = _query(
        file_path,
        "SELECT first || last FROM t WHERE id = 3",
    )
    assert rows == [(None,)]


def test_substr_uses_one_indexed_positions(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_substr.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT SUBSTR('hello', 2, 3), SUBSTR('hello', 2), SUBSTR('hello', 1, 2) FROM t WHERE id = 1",
    )
    assert rows == [("ell", "ello", "he")]


def test_year_month_day_functions(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_date_parts.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT YEAR(date_value), MONTH(date_value), DAY(date_value) FROM t WHERE id = 2",
    )
    assert rows == [(2023, 12, 5)]


def test_scalar_functions_with_join_results(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_join.xlsx"
    _create_scalar_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT t.id, CONCAT(UPPER(COALESCE(TRIM(t.name), 'unknown')), '-', LOWER(teams.team)) "
        "FROM t INNER JOIN teams ON t.id = teams.id ORDER BY t.id",
    )
    assert rows == [
        (1, "ALICE-eng"),
        (2, "BOB-ops"),
        (3, "UNKNOWN-sales"),
    ]


def test_scalar_functions_in_group_by_context(tmp_path: Path) -> None:
    file_path = tmp_path / "scalar_group_by.xlsx"
    _create_scalar_workbook(file_path)

    rows, description = _query(
        file_path,
        "SELECT UPPER(TRIM(name)) AS normalized_name, COUNT(*) "
        "FROM t WHERE name IS NOT NULL "
        "GROUP BY UPPER(TRIM(name)) "
        "ORDER BY normalized_name",
    )
    assert description == ["normalized_name", "COUNT(*)"]
    assert rows == [("ALICE", 1), ("BOB", 1)]
