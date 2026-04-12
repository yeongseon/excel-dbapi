from pathlib import Path

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
