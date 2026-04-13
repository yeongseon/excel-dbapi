from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.parser import parse_sql


def _create_compound_workbook(path: Path) -> None:
    workbook = Workbook()

    t1 = workbook.active
    assert t1 is not None
    t1.title = "t1"
    t1.append(["id", "name"])
    t1.append([1, "A"])
    t1.append([2, "B"])
    t1.append([3, None])
    t1.append([2, "B"])

    t2 = workbook.create_sheet("t2")
    t2.append(["id", "name"])
    t2.append([2, "B"])
    t2.append([3, None])
    t2.append([4, "D"])
    t2.append([4, "D"])

    t3 = workbook.create_sheet("t3")
    t3.append(["id", "name"])
    t3.append([4, "D"])
    t3.append([5, "E"])

    workbook.save(path)


def test_parser_accepts_compound_variants() -> None:
    parsed = parse_sql("SELECT id, name FROM t1 UNION SELECT id, name FROM t2")
    assert parsed["action"] == "COMPOUND"
    assert parsed["operators"] == ["UNION"]
    assert len(parsed["queries"]) == 2

    parsed = parse_sql("SELECT id FROM t1 UNION ALL SELECT id FROM t2")
    assert parsed["operators"] == ["UNION ALL"]

    parsed = parse_sql("SELECT id FROM t1 INTERSECT SELECT id FROM t2")
    assert parsed["operators"] == ["INTERSECT"]

    parsed = parse_sql("SELECT id FROM t1 EXCEPT SELECT id FROM t2")
    assert parsed["operators"] == ["EXCEPT"]

    parsed = parse_sql("SELECT id FROM t1 UNION SELECT id FROM t2 UNION SELECT id FROM t3")
    assert parsed["operators"] == ["UNION", "UNION"]
    assert len(parsed["queries"]) == 3


def test_union_basic(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_union.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM t1 UNION SELECT id, name FROM t2")
        assert cursor.fetchall() == [(1, "A"), (2, "B"), (3, None), (4, "D")]


def test_union_all_basic(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_union_all.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t1 UNION ALL SELECT id FROM t2")
        assert cursor.fetchall() == [(1,), (2,), (3,), (2,), (2,), (3,), (4,), (4,)]


def test_intersect_basic(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_intersect.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t1 INTERSECT SELECT id FROM t2")
        assert cursor.fetchall() == [(2,), (3,)]


def test_except_basic(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_except.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t1 EXCEPT SELECT id FROM t2")
        assert cursor.fetchall() == [(1,)]


def test_union_different_tables(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_diff_tables.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t1 UNION SELECT id FROM t3")
        assert cursor.fetchall() == [(1,), (2,), (3,), (4,), (5,)]


def test_union_with_where(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_where.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t1 WHERE id <= 2 UNION SELECT id FROM t2 WHERE id >= 3")
        assert cursor.fetchall() == [(1,), (2,), (3,), (4,)]


def test_union_column_count_mismatch(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_column_mismatch.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        with pytest.raises(ValueError, match="column counts"):
            conn.execute("SELECT id FROM t1 UNION SELECT id, name FROM t2")


def test_chained_union(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_chained_union.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t1 UNION SELECT id FROM t2 UNION SELECT id FROM t3")
        assert cursor.fetchall() == [(1,), (2,), (3,), (4,), (5,)]


def test_mixed_compound(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_mixed.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t1 UNION SELECT id FROM t2 EXCEPT SELECT id FROM t3")
        assert cursor.fetchall() == [(1,), (2,), (3,)]


def test_union_all_preserves_order(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_union_all_order.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id FROM t1 WHERE id <= 2 ORDER BY id DESC "
            "UNION ALL "
            "SELECT id FROM t2 WHERE id >= 4 ORDER BY id DESC"
        )
        assert cursor.fetchall() == [(2,), (2,), (1,), (4,), (4,)]


def test_intersect_empty_result(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_intersect_empty.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t1 INTERSECT SELECT id FROM t3")
        assert cursor.fetchall() == []


def test_except_all_removed(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_except_all_removed.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM t2 EXCEPT SELECT id FROM t2")
        assert cursor.fetchall() == []


def test_union_with_order_by(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_order_by.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id FROM t1 WHERE id <= 2 ORDER BY id DESC "
            "UNION "
            "SELECT id FROM t2 WHERE id >= 3 ORDER BY id DESC"
        )
        assert cursor.fetchall() == [(2,), (1,), (4,), (3,)]


def test_union_null_handling(tmp_path: Path) -> None:
    file_path = tmp_path / "compound_nulls.xlsx"
    _create_compound_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM t1 WHERE id = 3 UNION SELECT name FROM t2 WHERE id = 3")
        assert cursor.fetchall() == [(None,)]
