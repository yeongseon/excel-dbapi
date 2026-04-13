from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.exceptions import ProgrammingError


def _create_join_workbook(path: Path) -> None:
    workbook = Workbook()

    t1 = workbook.active
    assert t1 is not None
    t1.title = "t1"
    t1.append(["id", "val1"])
    t1.append([1, "a1"])
    t1.append([2, "a2"])
    t1.append([4, "a4"])

    t2 = workbook.create_sheet("t2")
    t2.append(["id", "val2"])
    t2.append([1, "b1"])
    t2.append([2, "b2"])
    t2.append([3, "b3"])

    t3 = workbook.create_sheet("t3")
    t3.append(["id", "val3"])
    t3.append([1, "c1"])
    t3.append([3, "c3"])
    t3.append([4, "c4"])

    t4 = workbook.create_sheet("t4")
    t4.append(["id", "val4"])
    t4.append([1, "d1"])
    t4.append([2, "d2"])
    t4.append([4, "d4"])

    workbook.save(path)


def _run_query(path: Path, query: str) -> list[tuple[object, ...]]:
    with ExcelConnection(str(path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(query)
        return cursor.fetchall()


def test_right_join_basic_matching_rows(tmp_path: Path) -> None:
    file_path = tmp_path / "right_join_basic.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, b.val2 FROM t1 a RIGHT JOIN t2 b ON a.id = b.id WHERE a.id IS NOT NULL ORDER BY b.id",
    )
    assert rows == [(1, "b1"), (2, "b2")]


def test_right_join_non_matching_right_rows_have_null_left(tmp_path: Path) -> None:
    file_path = tmp_path / "right_join_null_left.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, b.id FROM t1 a RIGHT JOIN t2 b ON a.id = b.id ORDER BY b.id",
    )
    assert rows == [(1, 1), (2, 2), (None, 3)]


def test_right_outer_join_syntax(tmp_path: Path) -> None:
    file_path = tmp_path / "right_outer_join.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, b.id FROM t1 a RIGHT OUTER JOIN t2 b ON a.id = b.id ORDER BY b.id",
    )
    assert rows == [(1, 1), (2, 2), (None, 3)]


def test_right_join_with_where_clause(tmp_path: Path) -> None:
    file_path = tmp_path / "right_join_where.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, b.id FROM t1 a RIGHT JOIN t2 b ON a.id = b.id WHERE b.id >= 2 ORDER BY b.id",
    )
    assert rows == [(2, 2), (None, 3)]


def test_right_join_with_order_by(tmp_path: Path) -> None:
    file_path = tmp_path / "right_join_order.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, b.id FROM t1 a RIGHT JOIN t2 b ON a.id = b.id ORDER BY b.id DESC",
    )
    assert rows == [(None, 3), (2, 2), (1, 1)]


def test_chained_inner_join_three_tables(tmp_path: Path) -> None:
    file_path = tmp_path / "chained_inner_three.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, b.val2, c.val3 FROM t1 a JOIN t2 b ON a.id = b.id JOIN t3 c ON b.id = c.id ORDER BY a.id",
    )
    assert rows == [(1, "b1", "c1")]


def test_chained_left_then_inner_join(tmp_path: Path) -> None:
    file_path = tmp_path / "chained_left_inner.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, b.val2, c.val3 FROM t1 a LEFT JOIN t2 b ON a.id = b.id JOIN t3 c ON b.id = c.id ORDER BY a.id",
    )
    assert rows == [(1, "b1", "c1")]


def test_chained_inner_then_left_join(tmp_path: Path) -> None:
    file_path = tmp_path / "chained_inner_left.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, c.val3 FROM t1 a JOIN t2 b ON a.id = b.id LEFT JOIN t3 c ON b.id = c.id ORDER BY a.id",
    )
    assert rows == [(1, "c1"), (2, None)]


def test_chained_three_joins_four_tables(tmp_path: Path) -> None:
    file_path = tmp_path / "chained_three_joins.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, d.val4 FROM t1 a JOIN t2 b ON a.id = b.id JOIN t3 c ON b.id = c.id JOIN t4 d ON c.id = d.id",
    )
    assert rows == [(1, "d1")]


def test_chained_join_where_across_multiple_tables(tmp_path: Path) -> None:
    file_path = tmp_path / "chained_where.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id LEFT JOIN t3 c ON b.id = c.id WHERE c.val3 IS NULL AND b.val2 = 'b2'",
    )
    assert rows == [(2,)]


def test_chained_join_with_order_by(tmp_path: Path) -> None:
    file_path = tmp_path / "chained_order_by.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, c.val3 FROM t1 a JOIN t2 b ON a.id = b.id LEFT JOIN t3 c ON b.id = c.id ORDER BY a.id DESC",
    )
    assert rows == [(2, None), (1, "c1")]


def test_chained_right_then_left_join(tmp_path: Path) -> None:
    file_path = tmp_path / "chained_right_left.xlsx"
    _create_join_workbook(file_path)

    rows = _run_query(
        file_path,
        "SELECT a.id, b.id, c.val3 FROM t1 a RIGHT JOIN t2 b ON a.id = b.id LEFT JOIN t3 c ON b.id = c.id ORDER BY b.id",
    )
    assert rows == [(1, 1, "c1"), (2, 2, None), (None, 3, "c3")]


def test_right_join_without_on_raises_error(tmp_path: Path) -> None:
    file_path = tmp_path / "right_join_no_on.xlsx"
    _create_join_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError, match="JOIN requires ON condition"):
            cursor.execute("SELECT a.id FROM t1 a RIGHT JOIN t2 b")


def test_chained_join_duplicate_alias_raises_error(tmp_path: Path) -> None:
    file_path = tmp_path / "chained_duplicate_alias.xlsx"
    _create_join_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        with pytest.raises(ProgrammingError, match="Ambiguous table reference"):
            cursor.execute(
                "SELECT a.id FROM t1 a JOIN t2 b ON a.id = b.id JOIN t3 b ON b.id = b.id"
            )


def test_chained_join_on_references_first_table(tmp_path: Path) -> None:
    """Third JOIN's ON clause references the first table, not the immediately previous one."""
    file_path = tmp_path / "chained_on_first.xlsx"
    _create_join_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        # t3 ON clause references a.id (first table), not b.id (immediate previous)
        cursor.execute(
            "SELECT a.id, b.val2, c.val3 FROM t1 a "
            "JOIN t2 b ON a.id = b.id "
            "LEFT JOIN t3 c ON a.id = c.id "
            "ORDER BY a.id"
        )
        rows = cursor.fetchall()
        # t1 has ids 1,2,4; t2 has 1,2,3; inner join gives 1,2
        # t3 has 1,3,4; left join on a.id=c.id for ids 1,2 gives:
        # id=1: c.val3='c1'; id=2: c.val3=None (no match)
        assert len(rows) == 2
        assert rows[0] == (1, "b1", "c1")
        assert rows[1] == (2, "b2", None)
