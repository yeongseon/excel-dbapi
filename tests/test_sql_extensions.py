from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.parser import parse_sql


def _create_extensions_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet.append(["id", "name", "score"])
    sheet.append([1, "Alice", 85])
    sheet.append([2, "Bob", 72])
    sheet.append([3, "Charlie", 95])
    sheet.append([4, "Diana", 88])
    sheet.append([5, None, 60])
    workbook.save(path)


def test_parse_where_in_values() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x IN (1, 2, 3)")
    where = parsed["where"]
    assert where["conditions"] == [
        {"column": "x", "operator": "IN", "value": (1, 2, 3)}
    ]


def test_parse_where_in_string_values() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x IN ('a', 'b')")
    condition = parsed["where"]["conditions"][0]
    assert condition["operator"] == "IN"
    assert condition["value"] == ("a", "b")


def test_parse_where_in_single_value() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x IN (1)")
    assert parsed["where"]["conditions"][0]["value"] == (1,)


def test_parse_where_in_rejects_empty_list() -> None:
    with pytest.raises(ValueError, match="IN clause cannot be empty"):
        parse_sql("SELECT * FROM t WHERE x IN ()")


def test_parse_where_between() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x BETWEEN 1 AND 10")
    where = parsed["where"]
    assert where["conditions"] == [
        {"column": "x", "operator": "BETWEEN", "value": (1, 10)}
    ]


def test_parse_where_like_percent_pattern() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x LIKE '%pattern%'")
    assert parsed["where"]["conditions"][0] == {
        "column": "x",
        "operator": "LIKE",
        "value": "%pattern%",
    }


def test_parse_where_like_underscore_pattern() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x LIKE 'test_'")
    assert parsed["where"]["conditions"][0]["value"] == "test_"


def test_parse_where_in_and_like_combined() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x IN (1,2) AND y LIKE '%foo%'")
    where = parsed["where"]
    assert where["conditions"][0] == {"column": "x", "operator": "IN", "value": (1, 2)}
    assert where["conditions"][1] == {
        "column": "y",
        "operator": "LIKE",
        "value": "%foo%",
    }
    assert where["conjunctions"] == ["AND"]


def test_parse_where_between_and_equals_combined() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x BETWEEN 1 AND 10 AND y = 'hello'")
    where = parsed["where"]
    assert where["conditions"][0] == {
        "column": "x",
        "operator": "BETWEEN",
        "value": (1, 10),
    }
    assert where["conditions"][1] == {"column": "y", "operator": "=", "value": "hello"}
    assert where["conjunctions"] == ["AND"]


def test_parse_where_in_param_binding() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x IN (?, ?)", (1, 2))
    assert parsed["where"]["conditions"][0]["value"] == (1, 2)


def test_parse_where_between_param_binding() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x BETWEEN ? AND ?", (3, 9))
    assert parsed["where"]["conditions"][0]["value"] == (3, 9)


def test_parse_where_like_param_binding() -> None:
    parsed = parse_sql("SELECT * FROM t WHERE x LIKE ?", ("ab%",))
    assert parsed["where"]["conditions"][0]["value"] == "ab%"


def test_execute_where_in_operator(tmp_path: Path) -> None:
    file_path = tmp_path / "extensions.xlsx"
    _create_extensions_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, name FROM Sheet1 WHERE name IN ('Alice', 'Bob') ORDER BY id"
        )
        assert cursor.fetchall() == [(1, "Alice"), (2, "Bob")]

        cursor.execute("SELECT id FROM Sheet1 WHERE name IN ('NonExistent')")
        assert cursor.fetchall() == []


def test_execute_where_between_operator(tmp_path: Path) -> None:
    file_path = tmp_path / "extensions.xlsx"
    _create_extensions_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM Sheet1 WHERE score BETWEEN 5 AND 15 ORDER BY id")
        assert cursor.fetchall() == []

        cursor.execute(
            "SELECT id FROM Sheet1 WHERE score BETWEEN 70 AND 90 ORDER BY id"
        )
        assert cursor.fetchall() == [(1,), (2,), (4,)]

        cursor.execute("SELECT id FROM Sheet1 WHERE score BETWEEN 100 AND 200")
        assert cursor.fetchall() == []


def test_execute_where_like_operator(tmp_path: Path) -> None:
    file_path = tmp_path / "extensions.xlsx"
    _create_extensions_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM Sheet1 WHERE name LIKE 'A%'")
        assert cursor.fetchall() == [(1,)]

        cursor.execute("SELECT id FROM Sheet1 WHERE name LIKE '%li%' ORDER BY id")
        assert cursor.fetchall() == [(1,), (3,)]

        cursor.execute("SELECT id FROM Sheet1 WHERE name LIKE 'Ali_e'")
        assert cursor.fetchall() == [(1,)]

        cursor.execute("SELECT id FROM Sheet1 WHERE name LIKE '%z%'")
        assert cursor.fetchall() == []


def test_execute_where_null_handling_for_new_operators(tmp_path: Path) -> None:
    file_path = tmp_path / "extensions.xlsx"
    _create_extensions_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id FROM Sheet1 WHERE name IN ('Alice', 'Bob', 'Charlie', 'Diana')"
        )
        assert cursor.fetchall() == [(1,), (2,), (3,), (4,)]

        cursor.execute("SELECT id FROM Sheet1 WHERE name LIKE '%' ORDER BY id")
        assert cursor.fetchall() == [(1,), (2,), (3,), (4,)]

        cursor.execute(
            "SELECT id FROM Sheet1 WHERE score BETWEEN 50 AND 90 ORDER BY id"
        )
        assert cursor.fetchall() == [(1,), (2,), (4,), (5,)]


def test_execute_where_with_order_by_and_limit(tmp_path: Path) -> None:
    file_path = tmp_path / "extensions.xlsx"
    _create_extensions_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, name FROM Sheet1 WHERE score BETWEEN 70 AND 95 ORDER BY id DESC LIMIT 2"
        )
        assert cursor.fetchall() == [(4, "Diana"), (3, "Charlie")]
