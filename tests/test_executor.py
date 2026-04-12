from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.engines.openpyxl.backend import OpenpyxlBackend
from excel_dbapi.executor import SharedExecutor
from excel_dbapi.parser import parse_sql


def test_executor_select():
    engine = OpenpyxlBackend("tests/data/sample.xlsx")
    parsed = parse_sql("SELECT * FROM Sheet1")
    results = SharedExecutor(engine).execute(parsed)
    assert isinstance(results.rows, list)
    assert isinstance(results.rows[0], tuple)


def test_executor_select_with_where():
    engine = OpenpyxlBackend("tests/data/sample.xlsx")
    parsed = parse_sql("SELECT * FROM Sheet1 WHERE id = 1")
    results = SharedExecutor(engine).execute(parsed)

    assert isinstance(results.rows, list)
    assert len(results.rows) == 1
    assert results.rows[0][0] == 1


def _create_select_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet.append(["id", "name", "score", "tags"])
    sheet.append([1, "A", 10, None])
    sheet.append([2, "A", 10, None])
    sheet.append([3, "B", None, None])
    sheet.append([4, "C", 30, "x"])
    sheet.append([5, "B", None, None])
    workbook.save(path)


def _create_empty_select_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet.append(["id", "name", "score"])
    workbook.save(path)


def _create_users_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "users"
    sheet.append(["id", "name", "age"])
    sheet.append([1, "Alice", 30])
    sheet.append([2, "Bob", 25])
    sheet.append([3, "Alice", 35])
    sheet.append([4, "Charlie", 40])
    workbook.save(path)


def _create_users_admins_workbook(path: Path) -> None:
    workbook = Workbook()

    users = workbook.active
    assert users is not None
    users.title = "users"
    users.append(["id", "name"])
    users.append([1, "Alice"])
    users.append([2, "Bob"])
    users.append([3, "Charlie"])

    admins = workbook.create_sheet("admins")
    admins.append(["id", "role"])
    admins.append([1, "admin"])
    admins.append([3, "editor"])

    workbook.save(path)


def test_executor_distinct_removes_duplicates_and_preserves_order(tmp_path: Path):
    file_path = tmp_path / "distinct_order.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT DISTINCT name, score FROM Sheet1 ORDER BY id ASC")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [("A", 10), ("B", None), ("C", 30)]


def test_executor_distinct_with_where_order_by_and_limit(tmp_path: Path):
    file_path = tmp_path / "distinct_where_limit.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT DISTINCT name FROM Sheet1 WHERE score IS NULL ORDER BY id ASC LIMIT 1"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [("B",)]


def test_executor_distinct_on_select_star(tmp_path: Path):
    file_path = tmp_path / "distinct_star.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet.append(["id", "name"])
    sheet.append([1, "A"])
    sheet.append([1, "A"])
    sheet.append([2, "B"])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT DISTINCT * FROM Sheet1 ORDER BY id ASC")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(1, "A"), (2, "B")]


def test_executor_offset_variants(tmp_path: Path):
    file_path = tmp_path / "offset_cases.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))

    parsed = parse_sql("SELECT id FROM Sheet1 ORDER BY id ASC OFFSET 2")
    results = SharedExecutor(engine).execute(parsed)
    assert results.rows == [(3,), (4,), (5,)]

    parsed = parse_sql("SELECT id FROM Sheet1 ORDER BY id ASC LIMIT 2 OFFSET 2")
    results = SharedExecutor(engine).execute(parsed)
    assert results.rows == [(3,), (4,)]

    parsed = parse_sql("SELECT id FROM Sheet1 ORDER BY id ASC OFFSET 99")
    results = SharedExecutor(engine).execute(parsed)
    assert results.rows == []

    parsed = parse_sql("SELECT id FROM Sheet1 ORDER BY id ASC OFFSET 0")
    results = SharedExecutor(engine).execute(parsed)
    assert results.rows == [(1,), (2,), (3,), (4,), (5,)]


def test_executor_offset_with_where_and_distinct_limit(tmp_path: Path):
    file_path = tmp_path / "offset_where_distinct.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT DISTINCT name FROM Sheet1 WHERE id >= 2 ORDER BY id ASC LIMIT 2 OFFSET 1"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [("B",), ("C",)]


def test_executor_aggregate_count_star(tmp_path: Path):
    file_path = tmp_path / "aggregate_count.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT COUNT(*) FROM Sheet1")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(5,)]
    assert results.description[0][0] == "COUNT(*)"


def test_executor_aggregate_sum_avg_min_max(tmp_path: Path):
    file_path = tmp_path / "aggregate_numeric.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT SUM(score), AVG(score), MIN(score), MAX(score) FROM Sheet1")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(50.0, 50.0 / 3.0, 10.0, 30.0)]


def test_executor_group_by_count(tmp_path: Path):
    file_path = tmp_path / "group_by_count.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT name, COUNT(*) FROM Sheet1 GROUP BY name ORDER BY name ASC")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [("A", 2), ("B", 2), ("C", 1)]


def test_executor_group_by_having_sum(tmp_path: Path):
    file_path = tmp_path / "group_by_having.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT name, SUM(score) FROM Sheet1 GROUP BY name HAVING SUM(score) > 15 ORDER BY name ASC"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [("A", 20.0), ("C", 30.0)]


def test_executor_aggregate_empty_table(tmp_path: Path):
    file_path = tmp_path / "aggregate_empty.xlsx"
    _create_empty_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT COUNT(*), SUM(score), AVG(score), MIN(score), MAX(score) FROM Sheet1"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(0, None, None, None, None)]


def test_executor_count_column_excludes_nulls(tmp_path: Path):
    file_path = tmp_path / "aggregate_count_nulls.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT COUNT(*), COUNT(score) FROM Sheet1")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(5, 3)]


def test_executor_group_by_with_order_limit_offset(tmp_path: Path):
    file_path = tmp_path / "group_by_order_limit_offset.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT name, COUNT(*) FROM Sheet1 GROUP BY name ORDER BY name DESC LIMIT 1 OFFSET 1"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [("B", 2)]


def test_executor_distinct_with_group_by(tmp_path: Path):
    file_path = tmp_path / "group_by_distinct.xlsx"
    _create_select_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT DISTINCT name FROM Sheet1 GROUP BY name ORDER BY name ASC"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [("A",), ("B",), ("C",)]


def test_executor_group_by_having_count_not_in_select(tmp_path: Path):
    file_path = tmp_path / "users_having_count_not_selected.xlsx"
    _create_users_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT name FROM users GROUP BY name HAVING COUNT(*) > 1")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [("Alice",)]


def test_executor_group_by_having_group_column_not_in_select(tmp_path: Path):
    file_path = tmp_path / "users_having_group_column_not_selected.xlsx"
    _create_users_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT COUNT(*) FROM users GROUP BY name HAVING name = 'Alice'")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(2,)]


def test_having_rejects_non_grouped_column(tmp_path: Path):
    file_path = tmp_path / "users_having_non_grouped_column.xlsx"
    _create_users_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT COUNT(*) FROM users GROUP BY name HAVING age > 30")

    with pytest.raises(
        ValueError,
        match="in HAVING must be a GROUP BY column or aggregate function",
    ):
        SharedExecutor(engine).execute(parsed)


def test_executor_group_by_order_by_group_key_not_in_select(tmp_path: Path):
    file_path = tmp_path / "users_order_by_group_key_not_selected.xlsx"
    _create_users_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT COUNT(*) FROM users GROUP BY name ORDER BY name")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(2,), (1,), (1,)]


def test_having_rejects_aggregate_with_expression_arg(tmp_path: Path):
    file_path = tmp_path / "users_having_expression_aggregate.xlsx"
    _create_users_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT COUNT(*) FROM users GROUP BY name HAVING SUM(age+1) > 1")

    with pytest.raises(ValueError, match="Unsupported aggregate expression"):
        SharedExecutor(engine).execute(parsed)


def test_order_by_rejects_aggregate_with_expression_arg(tmp_path: Path):
    file_path = tmp_path / "users_order_by_expression_aggregate.xlsx"
    _create_users_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT name, COUNT(*) FROM users GROUP BY name ORDER BY SUM(age+1)")

    with pytest.raises(ValueError, match="Unsupported aggregate expression"):
        SharedExecutor(engine).execute(parsed)


def test_subquery_in_where(tmp_path: Path):
    file_path = tmp_path / "subquery_in_where.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT id, name FROM users WHERE id IN (SELECT id FROM admins)")
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(1, "Alice"), (3, "Charlie")]


def test_subquery_returns_empty(tmp_path: Path):
    file_path = tmp_path / "subquery_empty.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT id, name FROM users WHERE id IN (SELECT id FROM admins WHERE role = 'root')"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == []


def test_subquery_with_where(tmp_path: Path):
    file_path = tmp_path / "subquery_with_where.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT id, name FROM users WHERE id IN (SELECT id FROM admins WHERE role = 'admin')"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(1, "Alice")]


def test_subquery_reexecution_safe(tmp_path: Path):
    file_path = tmp_path / "subquery_reexec.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql("SELECT id, name FROM users WHERE id IN (SELECT id FROM admins)")
    executor = SharedExecutor(engine)

    result1 = executor.execute(parsed)
    result2 = executor.execute(parsed)

    assert result1.rows == result2.rows == [(1, "Alice"), (3, "Charlie")]


# --- JOIN executor tests ---


def test_executor_inner_join(tmp_path: Path):
    file_path = tmp_path / "inner_join.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, a.name, b.role FROM users a INNER JOIN admins b ON a.id = b.id"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(1, "Alice", "admin"), (3, "Charlie", "editor")]
    assert len(results.description) == 3
    assert results.description[0][0] == "a.id"
    assert results.description[1][0] == "a.name"
    assert results.description[2][0] == "b.role"


def test_executor_left_join(tmp_path: Path):
    file_path = tmp_path / "left_join.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, a.name, b.role FROM users a LEFT JOIN admins b ON a.id = b.id"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [
        (1, "Alice", "admin"),
        (2, "Bob", None),
        (3, "Charlie", "editor"),
    ]


def test_executor_join_with_where(tmp_path: Path):
    file_path = tmp_path / "join_where.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, a.name, b.role FROM users a INNER JOIN admins b ON a.id = b.id WHERE b.role = 'admin'"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(1, "Alice", "admin")]


def test_executor_join_with_order_by(tmp_path: Path):
    file_path = tmp_path / "join_order.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, b.role FROM users a INNER JOIN admins b ON a.id = b.id ORDER BY a.id DESC"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(3, "editor"), (1, "admin")]


def test_executor_join_with_limit_offset(tmp_path: Path):
    file_path = tmp_path / "join_limit.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, a.name FROM users a LEFT JOIN admins b ON a.id = b.id LIMIT 2 OFFSET 1"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(2, "Bob"), (3, "Charlie")]


def test_executor_left_join_no_match(tmp_path: Path):
    """LEFT JOIN where right side has no matches at all -> all right columns None."""
    file_path = tmp_path / "left_join_no_match.xlsx"
    workbook = Workbook()
    users = workbook.active
    assert users is not None
    users.title = "users"
    users.append(["id", "name"])
    users.append([10, "Xander"])
    users.append([20, "Yara"])

    admins = workbook.create_sheet("admins")
    admins.append(["id", "role"])
    admins.append([1, "admin"])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, a.name, b.role FROM users a LEFT JOIN admins b ON a.id = b.id"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(10, "Xander", None), (20, "Yara", None)]


def test_executor_join_multi_condition_on(tmp_path: Path):
    """ON with AND (two equality conditions)."""
    file_path = tmp_path / "join_multi_on.xlsx"
    workbook = Workbook()
    left = workbook.active
    assert left is not None
    left.title = "employees"
    left.append(["dept", "grade", "name"])
    left.append(["eng", "senior", "Alice"])
    left.append(["eng", "junior", "Bob"])
    left.append(["hr", "senior", "Charlie"])

    right = workbook.create_sheet("salaries")
    right.append(["dept", "grade", "salary"])
    right.append(["eng", "senior", 150])
    right.append(["eng", "junior", 100])
    right.append(["hr", "senior", 120])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.name, b.salary FROM employees a INNER JOIN salaries b "
        "ON a.dept = b.dept AND a.grade = b.grade ORDER BY b.salary DESC"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [
        ("Alice", 150),
        ("Charlie", 120),
        ("Bob", 100),
    ]


def test_executor_join_table_name_qualifiers(tmp_path: Path):
    """Use table names (not aliases) as qualifiers."""
    file_path = tmp_path / "join_table_names.xlsx"
    _create_users_admins_workbook(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT users.id, admins.role FROM users INNER JOIN admins ON users.id = admins.id"
    )
    results = SharedExecutor(engine).execute(parsed)

    assert results.rows == [(1, "admin"), (3, "editor")]


def test_executor_join_missing_right_sheet(tmp_path: Path):
    """Joining a non-existent sheet raises ValueError."""
    file_path = tmp_path / "join_missing.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "users"
    sheet.append(["id"])
    sheet.append([1])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, b.id FROM users a INNER JOIN nonexistent b ON a.id = b.id"
    )
    with pytest.raises(ValueError, match="not found"):
        SharedExecutor(engine).execute(parsed)


def test_executor_join_null_keys_do_not_match(tmp_path: Path):
    """NULL join keys must not match per SQL standard: NULL != NULL."""
    file_path = tmp_path / "join_null_keys.xlsx"
    workbook = Workbook()
    left = workbook.active
    assert left is not None
    left.title = "left_t"
    left.append(["id", "name"])
    left.append([1, "Alice"])
    left.append([None, "Bob"])  # NULL key
    left.append([3, "Charlie"])

    right = workbook.create_sheet("right_t")
    right.append(["id", "role"])
    right.append([1, "admin"])
    right.append([None, "ghost"])  # NULL key
    right.append([3, "editor"])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    # INNER JOIN: NULL keys should NOT match
    parsed = parse_sql(
        "SELECT a.id, a.name, b.role FROM left_t a INNER JOIN right_t b ON a.id = b.id"
    )
    results = SharedExecutor(engine).execute(parsed)
    # Bob should NOT appear — NULL != NULL
    assert results.rows == [(1, "Alice", "admin"), (3, "Charlie", "editor")]


def test_executor_left_join_null_keys_unmatched(tmp_path: Path):
    """LEFT JOIN with NULL keys: left row with NULL key gets NULL-filled right columns."""
    file_path = tmp_path / "left_join_null_keys.xlsx"
    workbook = Workbook()
    left = workbook.active
    assert left is not None
    left.title = "left_t"
    left.append(["id", "name"])
    left.append([1, "Alice"])
    left.append([None, "Bob"])

    right = workbook.create_sheet("right_t")
    right.append(["id", "role"])
    right.append([1, "admin"])
    right.append([None, "ghost"])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, a.name, b.role FROM left_t a LEFT JOIN right_t b ON a.id = b.id"
    )
    results = SharedExecutor(engine).execute(parsed)
    # Bob (NULL key) has no match → right columns are None
    assert results.rows == [(1, "Alice", "admin"), (None, "Bob", None)]


def test_executor_join_numeric_string_coercion(tmp_path: Path):
    """Join key '1' (string) should match 1 (int) via numeric coercion."""
    file_path = tmp_path / "join_coercion.xlsx"
    workbook = Workbook()
    left = workbook.active
    assert left is not None
    left.title = "left_t"
    left.append(["id", "name"])
    left.append(["1", "Alice"])  # string '1'
    left.append(["2", "Bob"])

    right = workbook.create_sheet("right_t")
    right.append(["id", "role"])
    right.append([1, "admin"])  # int 1
    right.append([2, "editor"])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, a.name, b.role FROM left_t a INNER JOIN right_t b ON a.id = b.id"
    )
    results = SharedExecutor(engine).execute(parsed)
    assert len(results.rows) == 2
    assert results.rows[0][1] == "Alice"
    assert results.rows[0][2] == "admin"


def test_executor_join_unknown_select_column(tmp_path: Path):
    """SELECT referencing non-existent column in JOIN query raises ValueError."""
    file_path = tmp_path / "join_bad_col.xlsx"
    workbook = Workbook()
    left = workbook.active
    assert left is not None
    left.title = "users"
    left.append(["id", "name"])
    left.append([1, "Alice"])

    right = workbook.create_sheet("admins")
    right.append(["id", "role"])
    right.append([1, "admin"])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    # a.nonexistent doesn't exist in users sheet
    parsed = parse_sql(
        "SELECT a.nonexistent FROM users a INNER JOIN admins b ON a.id = b.id"
    )
    with pytest.raises(ValueError, match="Unknown column"):
        SharedExecutor(engine).execute(parsed)


def test_executor_join_unknown_on_column(tmp_path: Path):
    """ON referencing non-existent column raises ValueError."""
    file_path = tmp_path / "join_bad_on.xlsx"
    workbook = Workbook()
    left = workbook.active
    assert left is not None
    left.title = "users"
    left.append(["id", "name"])
    left.append([1, "Alice"])

    right = workbook.create_sheet("admins")
    right.append(["id", "role"])
    right.append([1, "admin"])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    parsed = parse_sql(
        "SELECT a.id, b.id FROM users a INNER JOIN admins b ON a.id = b.nonexistent"
    )
    with pytest.raises(ValueError, match="Unknown column"):
        SharedExecutor(engine).execute(parsed)


def test_executor_join_with_as_alias(tmp_path: Path):
    """JOIN with AS keyword in alias (SQLAlchemy compatibility)."""
    file_path = tmp_path / "join_as_alias.xlsx"
    workbook = Workbook()
    left = workbook.active
    assert left is not None
    left.title = "users"
    left.append(["id", "name"])
    left.append([1, "Alice"])
    left.append([2, "Bob"])

    right = workbook.create_sheet("admins")
    right.append(["id", "role"])
    right.append([1, "admin"])
    workbook.save(file_path)

    engine = OpenpyxlBackend(str(file_path))
    # Use 'AS' keyword in both FROM and JOIN aliases
    parsed = parse_sql(
        "SELECT a.id, a.name, b.role FROM users AS a INNER JOIN admins AS b ON a.id = b.id"
    )
    results = SharedExecutor(engine).execute(parsed)
    assert results.rows == [(1, "Alice", "admin")]
