from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.parser import parse_sql


def _create_round13_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "t"
    sheet.append(["a", "b", "val", "grp"])
    sheet.append([1, 1, 6, "x"])
    sheet.append([2, 3, 5, "x"])
    sheet.append([1, 1, 7, "y"])
    sheet.append([5, 5, 1, "y"])
    workbook.save(path)


def test_where_column_to_column_comparison_works_on_rhs_identifier(
    tmp_path: Path,
) -> None:
    file_path = tmp_path / "round13_where_column_rhs.xlsx"
    _create_round13_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT a, b FROM t WHERE a = b ORDER BY a")
        assert cursor.fetchall() == [(1, 1), (1, 1), (5, 5)]


def test_where_reversed_operands_resolve_rhs_identifier_as_column(
    tmp_path: Path,
) -> None:
    file_path = tmp_path / "round13_where_reversed_operands.xlsx"
    _create_round13_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT a FROM t WHERE 1 = a ORDER BY b")
        assert cursor.fetchall() == [(1,), (1,)]


def test_where_quoted_literal_still_treated_as_literal(tmp_path: Path) -> None:
    file_path = tmp_path / "round13_where_quoted_literal.xlsx"
    _create_round13_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT a FROM t WHERE a = 'b'")
        assert cursor.fetchall() == []


def test_having_aggregate_collected_from_rhs(tmp_path: Path) -> None:
    file_path = tmp_path / "round13_having_rhs_aggregate.xlsx"
    _create_round13_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT grp, SUM(val) FROM t GROUP BY grp HAVING 10 < SUM(val) ORDER BY grp"
        )
        assert cursor.fetchall() == [("x", 11.0)]


def test_having_aggregate_collection_handles_and_tree(tmp_path: Path) -> None:
    file_path = tmp_path / "round13_having_and_tree.xlsx"
    _create_round13_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT grp FROM t GROUP BY grp HAVING SUM(a) > 0 AND COUNT(*) > 1 ORDER BY grp"
        )
        assert cursor.fetchall() == [("x",), ("y",)]


def test_having_aggregate_collection_handles_not_wrapper(tmp_path: Path) -> None:
    file_path = tmp_path / "round13_having_not_wrapper.xlsx"
    _create_round13_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT grp FROM t GROUP BY grp HAVING NOT (SUM(val) > 10) ORDER BY grp"
        )
        assert cursor.fetchall() == [("y",)]


def test_parser_accepts_quoted_table_identifier_for_create_table() -> None:
    parsed = parse_sql('CREATE TABLE "Sales 2024" (id INTEGER, amount REAL)')
    assert parsed["table"] == "Sales 2024"


def test_select_from_quoted_table_name_with_space(tmp_path: Path) -> None:
    file_path = tmp_path / "round13_quoted_table_select.xlsx"

    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sales 2024"
    sheet.append(["id", "amount"])
    sheet.append([1, 100])
    workbook.save(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM "Sales 2024"')
        assert cursor.fetchall() == [(1,)]
