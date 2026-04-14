from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import ProgrammingError


def _create_round12_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "t"
    sheet.append(["id", "txt", "col"])
    sheet.append([1, "alpha", 1])
    sheet.append([2, "beta", 2])
    sheet.append([3, "gamma", 3])
    workbook.save(path)


def test_update_where_detection_ignores_where_inside_string_literal(
    tmp_path: Path,
) -> None:
    file_path = tmp_path / "round12_update_where_string.xlsx"
    _create_round12_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE t SET txt = 'value WHERE other' WHERE id = 1")
        cursor.execute("SELECT txt FROM t WHERE id = 1")
        assert cursor.fetchone() == ("value WHERE other",)


def test_not_in_with_null_candidates_returns_unknown_in_where(tmp_path: Path) -> None:
    file_path = tmp_path / "round12_not_in_null.xlsx"
    _create_round12_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t WHERE col NOT IN (2, NULL) ORDER BY id")
        assert cursor.fetchall() == []


def test_not_in_unknown_propagates_through_not_and_or(tmp_path: Path) -> None:
    file_path = tmp_path / "round12_not_in_compound_logic.xlsx"
    _create_round12_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()

        cursor.execute("SELECT id FROM t WHERE NOT (col NOT IN (2, NULL)) ORDER BY id")
        assert cursor.fetchall() == [(2,)]

        cursor.execute(
            "SELECT id FROM t "
            "WHERE NOT ((col NOT IN (2, NULL)) AND col != 3) ORDER BY id"
        )
        assert cursor.fetchall() == [(2,), (3,)]

        cursor.execute(
            "SELECT id FROM t WHERE NOT ((col NOT IN (2, NULL)) OR col = 3) ORDER BY id"
        )
        assert cursor.fetchall() == [(2,)]


def test_negative_limit_is_rejected_for_select_and_compound(tmp_path: Path) -> None:
    file_path = tmp_path / "round12_negative_limit.xlsx"
    _create_round12_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()

        with pytest.raises(
            ProgrammingError, match="LIMIT must be a non-negative integer"
        ):
            cursor.execute("SELECT id FROM t LIMIT -1")

        with pytest.raises(
            ProgrammingError, match="LIMIT must be a non-negative integer"
        ):
            cursor.execute("SELECT id FROM t UNION SELECT id FROM t LIMIT -1")


def test_negative_offset_is_rejected_for_select_and_compound(tmp_path: Path) -> None:
    file_path = tmp_path / "round12_negative_offset.xlsx"
    _create_round12_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()

        with pytest.raises(
            ProgrammingError,
            match="OFFSET must be a non-negative integer",
        ):
            cursor.execute("SELECT id FROM t OFFSET -1")

        with pytest.raises(
            ProgrammingError,
            match="OFFSET must be a non-negative integer",
        ):
            cursor.execute("SELECT id FROM t UNION SELECT id FROM t OFFSET -1")
