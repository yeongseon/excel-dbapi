from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import ProgrammingError


def _create_distinct_join_workbook(path: Path) -> None:
    workbook = Workbook()

    t1 = workbook.active
    assert t1 is not None
    t1.title = "t1"
    t1.append(["id", "col", "grp"])
    t1.append([1, "x", "g1"])
    t1.append([2, "x", "g1"])
    t1.append([3, "y", "g2"])
    t1.append([4, "y", "g2"])

    t2 = workbook.create_sheet("t2")
    t2.append(["id", "tag", "score"])
    t2.append([1, "p", 10])
    t2.append([1, "q", 20])
    t2.append([2, "p", 30])
    t2.append([2, "p", 30])
    t2.append([3, "p", 40])
    t2.append([5, "z", 50])

    workbook.save(path)


def _run_query(path: Path, query: str) -> list[tuple[object, ...]]:
    with ExcelConnection(str(path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(query)
        return cursor.fetchall()


def test_distinct_with_inner_join(tmp_path: Path) -> None:
    file_path = tmp_path / "distinct_inner_join.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT DISTINCT a.col FROM t1 a JOIN t2 b ON a.id = b.id ORDER BY a.col",
    )
    assert rows == [("x",), ("y",)]


def test_distinct_with_left_join(tmp_path: Path) -> None:
    file_path = tmp_path / "distinct_left_join.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT DISTINCT a.col FROM t1 a LEFT JOIN t2 b ON a.id = b.id ORDER BY a.col",
    )
    assert rows == [("x",), ("y",)]


def test_distinct_star_with_join(tmp_path: Path) -> None:
    file_path = tmp_path / "distinct_star_join.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT DISTINCT * FROM t1 a JOIN t2 b ON a.id = b.id ORDER BY a.id, b.tag",
    )
    assert rows == [
        (1, "x", "g1", 1, "p", 10),
        (1, "x", "g1", 1, "q", 20),
        (2, "x", "g1", 2, "p", 30),
        (3, "y", "g2", 3, "p", 40),
    ]


def test_distinct_specific_columns_with_join(tmp_path: Path) -> None:
    file_path = tmp_path / "distinct_columns_join.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT DISTINCT a.col, b.tag FROM t1 a JOIN t2 b ON a.id = b.id ORDER BY a.col, b.tag",
    )
    assert rows == [("x", "p"), ("x", "q"), ("y", "p")]


def test_count_distinct_qualified_with_join(tmp_path: Path) -> None:
    file_path = tmp_path / "count_distinct_qualified_join.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT COUNT(DISTINCT a.col) FROM t1 a JOIN t2 b ON a.id = b.id",
    )
    assert rows == [(2,)]


def test_count_distinct_qualified_without_join(tmp_path: Path) -> None:
    file_path = tmp_path / "count_distinct_qualified_single.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(file_path, "SELECT COUNT(DISTINCT a.col) FROM t1 a")
    assert rows == [(2,)]


def test_distinct_join_order_by_selected_column(tmp_path: Path) -> None:
    file_path = tmp_path / "distinct_join_order_by_selected.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT DISTINCT a.col FROM t1 a JOIN t2 b ON a.id = b.id ORDER BY a.col DESC",
    )
    assert rows == [("y",), ("x",)]


def test_distinct_join_order_by_non_selected_column_rejected(tmp_path: Path) -> None:
    file_path = tmp_path / "distinct_join_order_by_non_selected.xlsx"
    _create_distinct_join_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(
            ProgrammingError,
        match="ORDER BY columns must appear in SELECT list when using DISTINCT",
        ):
            cursor.execute(
                "SELECT DISTINCT a.col FROM t1 a JOIN t2 b ON a.id = b.id ORDER BY b.tag"
            )


def test_distinct_with_group_by_join(tmp_path: Path) -> None:
    file_path = tmp_path / "distinct_group_by_join.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT DISTINCT a.col, COUNT(*) FROM t1 a JOIN t2 b ON a.id = b.id GROUP BY a.col ORDER BY a.col",
    )
    assert rows == [("x", 4), ("y", 1)]


def test_distinct_join_all_rows_identical_after_projection(tmp_path: Path) -> None:
    file_path = tmp_path / "distinct_join_identical_projection.xlsx"
    _create_distinct_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT DISTINCT a.col FROM t1 a JOIN t2 b ON a.id = b.id WHERE a.id = 2",
    )
    assert rows == [("x",)]
