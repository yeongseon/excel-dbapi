from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path
from typing import cast

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import pytest

from excel_dbapi import connect


@pytest.fixture
def make_workbook(tmp_path: Path) -> Callable[..., Path]:
    def _make(
        file_name: str = "sql_spec.xlsx", include_join_sheet: bool = False
    ) -> Path:
        file_path = tmp_path / file_name
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Sheet1"
        sheet.append(["id", "name", "score", "created_at"])
        sheet.append([1, "Alice", 85, "2024-01-15"])
        sheet.append([2, "Bob", 72, "2024-02-20"])
        sheet.append([3, "Carol", 95, "2024-03-10"])

        if include_join_sheet:
            join_sheet = cast(Worksheet, workbook.create_sheet("Sheet2"))
            join_sheet.append(["user_id", "team"])
            join_sheet.append([1, "Alpha"])
            join_sheet.append([3, "Beta"])
            join_sheet.append([4, "Gamma"])

        workbook.save(file_path)
        return file_path

    return _make


def _run_select(
    file_path: Path, sql: str, params: Sequence[object] | None = None
) -> list[tuple[object, ...]]:
    with connect(str(file_path)) as conn:
        if params is None:
            return conn.execute(sql).rows
        return conn.execute(sql, params).rows


@pytest.mark.parametrize(
    ("sql", "expected"),
    [
        (
            "SELECT * FROM Sheet1 ORDER BY id",
            [
                (1, "Alice", 85, "2024-01-15"),
                (2, "Bob", 72, "2024-02-20"),
                (3, "Carol", 95, "2024-03-10"),
            ],
        ),
        (
            "SELECT id, name FROM Sheet1 ORDER BY id",
            [(1, "Alice"), (2, "Bob"), (3, "Carol")],
        ),
        (
            "SELECT name AS person_name, score AS points FROM Sheet1 ORDER BY id",
            [("Alice", 85), ("Bob", 72), ("Carol", 95)],
        ),
        (
            "SELECT DISTINCT score FROM Sheet1 ORDER BY score",
            [(72,), (85,), (95,)],
        ),
        (
            "SELECT id FROM Sheet1 ORDER BY id LIMIT 1 OFFSET 1",
            [(2,)],
        ),
    ],
)
def test_select_basics(
    make_workbook: Callable[..., Path],
    sql: str,
    expected: list[tuple[object, ...]],
) -> None:
    file_path = make_workbook("select_basics.xlsx")
    assert _run_select(file_path, sql) == expected


@pytest.mark.parametrize(
    ("setup_sql", "sql", "expected_ids"),
    [
        ((), "SELECT id FROM Sheet1 WHERE id = 2 ORDER BY id", [2]),
        ((), "SELECT id FROM Sheet1 WHERE id != 2 ORDER BY id", [1, 3]),
        ((), "SELECT id FROM Sheet1 WHERE id <> 2 ORDER BY id", [1, 3]),
        ((), "SELECT id FROM Sheet1 WHERE score > 80 ORDER BY id", [1, 3]),
        ((), "SELECT id FROM Sheet1 WHERE score >= 85 ORDER BY id", [1, 3]),
        ((), "SELECT id FROM Sheet1 WHERE score < 85 ORDER BY id", [2]),
        ((), "SELECT id FROM Sheet1 WHERE score <= 85 ORDER BY id", [1, 2]),
        (
            ("UPDATE Sheet1 SET created_at = NULL WHERE id = 2",),
            "SELECT id FROM Sheet1 WHERE created_at IS NULL ORDER BY id",
            [2],
        ),
        (
            ("UPDATE Sheet1 SET created_at = NULL WHERE id = 2",),
            "SELECT id FROM Sheet1 WHERE created_at IS NOT NULL ORDER BY id",
            [1, 3],
        ),
        ((), "SELECT id FROM Sheet1 WHERE id IN (1, 3) ORDER BY id", [1, 3]),
        ((), "SELECT id FROM Sheet1 WHERE id NOT IN (1, 3) ORDER BY id", [2]),
        ((), "SELECT id FROM Sheet1 WHERE score BETWEEN 80 AND 90 ORDER BY id", [1]),
        (
            (),
            "SELECT id FROM Sheet1 WHERE score NOT BETWEEN 80 AND 90 ORDER BY id",
            [2, 3],
        ),
        ((), "SELECT id FROM Sheet1 WHERE name LIKE 'A%' ORDER BY id", [1]),
        ((), "SELECT id FROM Sheet1 WHERE name ILIKE 'a%' ORDER BY id", [1]),
        ((), "SELECT id FROM Sheet1 WHERE name NOT LIKE 'A%' ORDER BY id", [2, 3]),
        ((), "SELECT id FROM Sheet1 WHERE id > 1 AND score > 80 ORDER BY id", [3]),
        ((), "SELECT id FROM Sheet1 WHERE id = 1 OR id = 3 ORDER BY id", [1, 3]),
        ((), "SELECT id FROM Sheet1 WHERE NOT (score > 80) ORDER BY id", [2]),
    ],
)
def test_where_operators(
    make_workbook: Callable[..., Path],
    setup_sql: tuple[str, ...],
    sql: str,
    expected_ids: list[int],
) -> None:
    file_path = make_workbook("where_ops.xlsx")

    with connect(str(file_path)) as conn:
        for statement in setup_sql:
            _ = conn.execute(statement)
        rows = conn.execute(sql).rows

    assert [row[0] for row in rows] == expected_ids


@pytest.mark.parametrize(
    ("setup_sql", "sql", "expected"),
    [
        ((), "SELECT id FROM Sheet1 ORDER BY score ASC, id ASC", [(2,), (1,), (3,)]),
        ((), "SELECT id FROM Sheet1 ORDER BY score DESC, id DESC", [(3,), (1,), (2,)]),
        (
            ("INSERT INTO Sheet1 VALUES (4, 'Aaron', 85, '2024-04-01')",),
            "SELECT id FROM Sheet1 ORDER BY score DESC, name ASC",
            [(3,), (4,), (1,), (2,)],
        ),
    ],
)
def test_order_by(
    make_workbook: Callable[..., Path],
    setup_sql: tuple[str, ...],
    sql: str,
    expected: list[tuple[object, ...]],
) -> None:
    file_path = make_workbook("order_by.xlsx")

    with connect(str(file_path)) as conn:
        for statement in setup_sql:
            _ = conn.execute(statement)
        rows = conn.execute(sql).rows

    assert rows == expected


@pytest.mark.parametrize(
    ("setup_sql", "sql", "expected"),
    [
        ((), "SELECT COUNT(*) FROM Sheet1", [(3,)]),
        ((), "SELECT COUNT(created_at) FROM Sheet1", [(3,)]),
        (
            ("INSERT INTO Sheet1 VALUES (4, 'Alice', 88, '2024-04-01')",),
            "SELECT COUNT(DISTINCT name) FROM Sheet1",
            [(3,)],
        ),
        ((), "SELECT SUM(score) FROM Sheet1", [(252,)]),
        ((), "SELECT AVG(score) FROM Sheet1", [(84.0,)]),
        ((), "SELECT MIN(score) FROM Sheet1", [(72,)]),
        ((), "SELECT MAX(score) FROM Sheet1", [(95,)]),
    ],
)
def test_aggregates(
    make_workbook: Callable[..., Path],
    setup_sql: tuple[str, ...],
    sql: str,
    expected: list[tuple[object, ...]],
) -> None:
    file_path = make_workbook("aggregates.xlsx")

    with connect(str(file_path)) as conn:
        for statement in setup_sql:
            _ = conn.execute(statement)
        rows = conn.execute(sql).rows

    assert rows == expected


def test_group_by_and_having(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("group_having.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("INSERT INTO Sheet1 VALUES (4, 'Alice', 88, '2024-04-01')")
        rows = conn.execute(
            "SELECT name, COUNT(*) FROM Sheet1 GROUP BY name HAVING COUNT(*) > 1 ORDER BY name"
        ).rows

    assert rows == [("Alice", 2)]


@pytest.mark.parametrize(
    ("sql", "expected"),
    [
        ("SELECT UPPER(name) FROM Sheet1 WHERE id = 1", [("ALICE",)]),
        ("SELECT LOWER(name) FROM Sheet1 WHERE id = 1", [("alice",)]),
        ("SELECT LENGTH(name) FROM Sheet1 WHERE id = 1", [(5,)]),
        ("SELECT TRIM('  hello  ') FROM Sheet1 LIMIT 1", [("hello",)]),
        ("SELECT SUBSTR(name, 1, 3) FROM Sheet1 WHERE id = 1", [("Ali",)]),
        ("SELECT COALESCE(NULL, name, 'x') FROM Sheet1 WHERE id = 1", [("Alice",)]),
        ("SELECT NULLIF(name, 'Alice') FROM Sheet1 WHERE id = 1", [(None,)]),
        ("SELECT CONCAT(name, '-ok') FROM Sheet1 WHERE id = 1", [("Alice-ok",)]),
        ("SELECT ABS(-7) FROM Sheet1 LIMIT 1", [(7,)]),
        ("SELECT ROUND(12.34, 1) FROM Sheet1 LIMIT 1", [(12.3,)]),
        ("SELECT REPLACE(name, 'li', 'xx') FROM Sheet1 WHERE id = 1", [("Axxce",)]),
    ],
)
def test_scalar_functions(
    make_workbook: Callable[..., Path],
    sql: str,
    expected: list[tuple[object, ...]],
) -> None:
    file_path = make_workbook("scalar_functions.xlsx")
    assert _run_select(file_path, sql) == expected


@pytest.mark.parametrize(
    ("sql", "expected"),
    [
        (
            "SELECT s1.id, s2.team FROM Sheet1 AS s1 INNER JOIN Sheet2 AS s2 ON s1.id = s2.user_id ORDER BY s1.id",
            [(1, "Alpha"), (3, "Beta")],
        ),
        (
            "SELECT s1.id, s2.team FROM Sheet1 AS s1 LEFT JOIN Sheet2 AS s2 ON s1.id = s2.user_id ORDER BY s1.id",
            [(1, "Alpha"), (2, None), (3, "Beta")],
        ),
    ],
)
def test_joins(
    make_workbook: Callable[..., Path],
    sql: str,
    expected: list[tuple[object, ...]],
) -> None:
    file_path = make_workbook("joins.xlsx", include_join_sheet=True)
    assert _run_select(file_path, sql) == expected


def test_insert_single_row(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("insert_single.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("INSERT INTO Sheet1 VALUES (4, 'Dave', 88, '2024-04-01')")
        rows = conn.execute("SELECT name, score FROM Sheet1 WHERE id = 4").rows

    assert rows == [("Dave", 88)]


def test_insert_multi_row(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("insert_multi.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute(
            "INSERT INTO Sheet1 VALUES (4, 'Dan', 60, '2024-04-01'), (5, 'Eve', 91, '2024-04-02')"
        )
        rows = conn.execute("SELECT id FROM Sheet1 ORDER BY id").rows

    assert rows == [(1,), (2,), (3,), (4,), (5,)]


def test_update_with_where(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("update_where.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("UPDATE Sheet1 SET score = 100 WHERE id = 2")
        rows = conn.execute("SELECT score FROM Sheet1 WHERE id = 2").rows

    assert rows == [(100,)]


def test_delete_with_where(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("delete_where.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("DELETE FROM Sheet1 WHERE id = 1")
        rows = conn.execute("SELECT id FROM Sheet1 ORDER BY id").rows

    assert rows == [(2,), (3,)]


def test_create_and_drop_table(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("ddl_create_drop.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("CREATE TABLE TempSheet (id INTEGER, note TEXT)")
        _ = conn.execute("INSERT INTO TempSheet VALUES (1, 'ok')")
        first_rows = conn.execute("SELECT id, note FROM TempSheet").rows
        _ = conn.execute("DROP TABLE TempSheet")
        _ = conn.execute("CREATE TABLE TempSheet (id INTEGER, note TEXT)")
        second_rows = conn.execute("SELECT * FROM TempSheet").rows

    assert first_rows == [(1, "ok")]
    assert second_rows == []


@pytest.mark.parametrize(
    ("sql", "expected"),
    [
        (
            "SELECT score FROM Sheet1 WHERE id IN (1, 2) UNION SELECT score FROM Sheet1 WHERE id IN (2, 3) ORDER BY score",
            [(72,), (85,), (95,)],
        ),
        (
            "SELECT score FROM Sheet1 WHERE id IN (1, 2) UNION ALL SELECT score FROM Sheet1 WHERE id IN (2, 3) ORDER BY score",
            [(72,), (72,), (85,), (95,)],
        ),
    ],
)
def test_set_operations(
    make_workbook: Callable[..., Path],
    sql: str,
    expected: list[tuple[object, ...]],
) -> None:
    file_path = make_workbook("set_ops.xlsx")
    assert _run_select(file_path, sql) == expected


def test_subquery_in_where(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("subquery_in.xlsx")
    rows = _run_select(
        file_path,
        "SELECT id FROM Sheet1 WHERE id IN (SELECT id FROM Sheet1 WHERE score >= 85) ORDER BY id",
    )
    assert rows == [(1,), (3,)]


@pytest.mark.parametrize(
    ("sql", "expected"),
    [
        ("SELECT CAST(score AS INTEGER) FROM Sheet1 WHERE id = 1", [(85,)]),
        ("SELECT CAST(score AS TEXT) FROM Sheet1 WHERE id = 1", [("85",)]),
    ],
)
def test_cast_expressions(
    make_workbook: Callable[..., Path],
    sql: str,
    expected: list[tuple[object, ...]],
) -> None:
    file_path = make_workbook("cast.xlsx")
    assert _run_select(file_path, sql) == expected


def test_parameter_binding_with_qmark(
    make_workbook: Callable[..., Path],
) -> None:
    file_path = make_workbook("params.xlsx")

    rows = _run_select(
        file_path,
        "SELECT name FROM Sheet1 WHERE score >= ? AND name LIKE ? ORDER BY id LIMIT ? OFFSET ?",
        (80, "%o%", 1, 0),
    )

    assert rows == [("Carol",)]


# ── Additional stable feature tests (Oracle review round) ──────────────


def test_right_join(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("right_join.xlsx", include_join_sheet=True)
    rows = _run_select(
        file_path,
        "SELECT s1.id, s2.team FROM Sheet1 AS s1 RIGHT JOIN Sheet2 AS s2 ON s1.id = s2.user_id ORDER BY s2.user_id",
    )
    assert rows == [(1, "Alpha"), (3, "Beta"), (None, "Gamma")]


def test_full_outer_join(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("full_join.xlsx", include_join_sheet=True)
    rows = _run_select(
        file_path,
        "SELECT s1.id, s2.team FROM Sheet1 AS s1 FULL OUTER JOIN Sheet2 AS s2 ON s1.id = s2.user_id ORDER BY s1.id",
    )
    # id=1 matches Alpha, id=2 has no match, id=3 matches Beta, user_id=4 (Gamma) has no match in Sheet1
    assert len(rows) == 4
    assert (1, "Alpha") in rows
    assert (2, None) in rows
    assert (3, "Beta") in rows
    assert (None, "Gamma") in rows

def test_cross_join(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("cross_join.xlsx", include_join_sheet=True)
    rows = _run_select(
        file_path,
        "SELECT s1.id, s2.team FROM Sheet1 AS s1 CROSS JOIN Sheet2 AS s2 ORDER BY s1.id, s2.team",
    )
    # 3 rows × 3 rows = 9 rows
    expected = [
        (1, "Alpha"), (1, "Beta"), (1, "Gamma"),
        (2, "Alpha"), (2, "Beta"), (2, "Gamma"),
        (3, "Alpha"), (3, "Beta"), (3, "Gamma"),
    ]
    assert rows == expected


def test_aggregate_filter(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("agg_filter.xlsx")
    rows = _run_select(
        file_path,
        "SELECT COUNT(*) FILTER (WHERE score > 80) FROM Sheet1",
    )
    assert rows == [(2,)]


def test_exists_subquery(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("exists.xlsx", include_join_sheet=True)
    rows = _run_select(
        file_path,
        "SELECT id FROM Sheet1 AS s1 WHERE EXISTS (SELECT 1 FROM Sheet2 AS s2 WHERE s2.user_id = s1.id) ORDER BY id",
    )
    assert rows == [(1,), (3,)]


def test_not_exists_subquery(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("not_exists.xlsx", include_join_sheet=True)
    rows = _run_select(
        file_path,
        "SELECT id FROM Sheet1 AS s1 WHERE NOT EXISTS (SELECT 1 FROM Sheet2 AS s2 WHERE s2.user_id = s1.id) ORDER BY id",
    )
    assert rows == [(2,)]


def test_intersect(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("intersect.xlsx")
    rows = _run_select(
        file_path,
        "SELECT id FROM Sheet1 WHERE id IN (1, 2) INTERSECT SELECT id FROM Sheet1 WHERE id IN (2, 3) ORDER BY id",
    )
    assert rows == [(2,)]


def test_except(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("except.xlsx")
    rows = _run_select(
        file_path,
        "SELECT id FROM Sheet1 WHERE id IN (1, 2) EXCEPT SELECT id FROM Sheet1 WHERE id IN (2, 3) ORDER BY id",
    )
    assert rows == [(1,)]


def test_insert_select(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("insert_select.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("CREATE TABLE Sheet3 (id, name, score, created_at)")
        _ = conn.execute("INSERT INTO Sheet3 (id, name, score, created_at) SELECT id, name, score, created_at FROM Sheet1 WHERE score >= 85")
        rows = conn.execute("SELECT id FROM Sheet3 ORDER BY id").rows

    assert rows == [(1,), (3,)]


def test_upsert_on_conflict_do_nothing(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("upsert_nothing.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("INSERT INTO Sheet1 VALUES (1, 'Duplicate', 0, '2024-01-01') ON CONFLICT (id) DO NOTHING")
        rows = conn.execute("SELECT name FROM Sheet1 WHERE id = 1").rows

    assert rows == [("Alice",)]


def test_upsert_on_conflict_do_update(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("upsert_update.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("INSERT INTO Sheet1 VALUES (1, 'Updated', 99, '2024-01-01') ON CONFLICT (id) DO UPDATE SET name = 'Updated', score = 99")
        rows = conn.execute("SELECT name, score FROM Sheet1 WHERE id = 1").rows

    assert rows == [("Updated", 99)]


def test_alter_table_add_column(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("alter_add.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("ALTER TABLE Sheet1 ADD COLUMN status TEXT")
        rows = conn.execute("SELECT status FROM Sheet1 WHERE id = 1").rows

    assert rows == [(None,)]


def test_alter_table_drop_column(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("alter_drop.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("ALTER TABLE Sheet1 DROP COLUMN created_at")
        result = conn.execute("SELECT * FROM Sheet1 WHERE id = 1")
        headers = [col[0] for col in result.description] if result.description else []

    assert "created_at" not in headers


def test_alter_table_rename_column(make_workbook: Callable[..., Path]) -> None:
    file_path = make_workbook("alter_rename.xlsx")

    with connect(str(file_path)) as conn:
        _ = conn.execute("ALTER TABLE Sheet1 RENAME COLUMN name TO full_name")
        rows = conn.execute("SELECT full_name FROM Sheet1 WHERE id = 1").rows

    assert rows == [("Alice",)]
