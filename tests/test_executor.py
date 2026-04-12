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
