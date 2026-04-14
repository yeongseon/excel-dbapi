from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import ProgrammingError


def _create_round11_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet.append(["id", "name"])
    sheet.append([1, "Alice"])
    sheet.append([2, "Bob"])

    table = workbook.create_sheet("t")
    table.append(["id", "a", "b", "c"])
    table.append([1, 10, 0, None])
    table.append([2, "alice", 0, None])

    workbook.save(path)


def test_executemany_failure_clears_prior_select_results(tmp_path: Path) -> None:
    file_path = tmp_path / "round11_cursor_state.xlsx"
    _create_round11_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=False) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM Sheet1 ORDER BY id")
        assert cursor.fetchone() == (1, "Alice")

        with pytest.raises(ProgrammingError):
            cursor.executemany(
                "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
                [(3, "Cara"), (4,)],
            )

        with pytest.raises(ProgrammingError, match="No result set"):
            cursor.fetchone()


def test_update_set_supports_column_arithmetic_cast_and_function(
    tmp_path: Path,
) -> None:
    file_path = tmp_path / "round11_update_expressions.xlsx"
    _create_round11_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()

        cursor.execute("UPDATE t SET b = a WHERE id = 1")
        cursor.execute("UPDATE t SET b = a + 1 WHERE id = 1")
        cursor.execute("UPDATE t SET b = CAST(a AS TEXT) WHERE id = 1")
        cursor.execute("UPDATE t SET c = UPPER(a) WHERE id = 2")

        cursor.execute("SELECT b, c FROM t WHERE id = 1")
        assert cursor.fetchone() == ("10", None)
        cursor.execute("SELECT b, c FROM t WHERE id = 2")
        assert cursor.fetchone() == (0, "ALICE")


def test_upsert_update_set_supports_expression_nodes(tmp_path: Path) -> None:
    file_path = tmp_path / "round11_upsert_expressions.xlsx"
    _create_round11_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO t (id, a, b, c) VALUES (1, 9, 0, 'incoming') "
            "ON CONFLICT (id) DO UPDATE SET b = a + 1, c = CAST(excluded.a AS TEXT)"
        )

        cursor.execute("SELECT a, b, c FROM t WHERE id = 1")
        assert cursor.fetchone() == (10, 11, "9")


def test_cast_supports_datetime_and_boolean_with_edge_cases(tmp_path: Path) -> None:
    file_path = tmp_path / "round11_cast_types.xlsx"
    _create_round11_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT "
            "CAST('42' AS INTEGER), "
            "CAST('3.25' AS REAL), "
            "CAST('2024-01-15' AS DATE), "
            "CAST('2024-01-15T10:11:12Z' AS DATETIME), "
            "CAST('yes' AS BOOLEAN), "
            "CAST(0 AS BOOLEAN), "
            "CAST(NULL AS DATETIME), "
            "CAST(NULL AS BOOLEAN) "
            "FROM t WHERE id = 1"
        )
        row = cursor.fetchone()
        assert row == (
            42,
            3.25,
            date(2024, 1, 15),
            datetime(2024, 1, 15, 10, 11, 12),
            True,
            False,
            None,
            None,
        )

        with pytest.raises(ProgrammingError, match="Cannot cast value .* to BOOLEAN"):
            cursor.execute("SELECT CAST('maybe' AS BOOLEAN) FROM t")

        with pytest.raises(
            ProgrammingError, match="Cannot cast empty string to DATETIME"
        ):
            cursor.execute("SELECT CAST('' AS DATETIME) FROM t")


def test_abs_round_replace_scalar_functions_are_supported(tmp_path: Path) -> None:
    file_path = tmp_path / "round11_scalar_functions.xlsx"
    _create_round11_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT ABS(-3), ABS('-2.5'), ROUND(3.14159, 2), ROUND(2.6), "
            "REPLACE('banana', 'na', 'NA') FROM t WHERE id = 1"
        )
        assert cursor.fetchone() == (3, 2.5, 3.14, 3, "baNANA")
