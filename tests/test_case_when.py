from pathlib import Path

from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.parser import parse_sql


def _create_case_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "t"
    sheet.append(["id", "score", "tier", "label", "points"])
    sheet.append([1, 15, "gold", "x", 2])
    sheet.append([2, 8, "silver", None, 3])
    sheet.append([3, None, "bronze", "z", 4])
    sheet.append([4, 12, "gold", "w", 0])
    workbook.save(path)


def _query(
    path: Path,
    sql: str,
    params: tuple[object, ...] | None = None,
) -> tuple[list[tuple[object, ...]], list[str]]:
    with ExcelConnection(str(path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(sql, params)
        rows = cursor.fetchall()
        description = [str(item[0]) for item in cursor.description or []]
        return rows, description


def _execute(path: Path, sql: str, params: tuple[object, ...] | None = None) -> int:
    with ExcelConnection(str(path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, params)
        return cursor.rowcount


def test_parser_builds_searched_case_ast() -> None:
    parsed = parse_sql(
        "SELECT CASE WHEN score > 10 THEN 'high' ELSE 'low' END AS bucket FROM t"
    )
    alias_expr = parsed["columns"][0]
    assert alias_expr["type"] == "alias"
    case_expr = alias_expr["expression"]
    assert case_expr["type"] == "case"
    assert case_expr["mode"] == "searched"
    assert case_expr["value"] is None
    assert case_expr["else"]["type"] == "literal"


def test_parser_builds_simple_case_ast() -> None:
    parsed = parse_sql(
        "SELECT CASE tier WHEN 'gold' THEN 1 ELSE 2 END FROM t"
    )
    case_expr = parsed["columns"][0]
    assert case_expr["type"] == "case"
    assert case_expr["mode"] == "simple"
    assert case_expr["value"] == "tier"
    assert case_expr["whens"][0]["match"]["type"] == "literal"


def test_parser_builds_case_in_update_set_and_where_operand() -> None:
    parsed = parse_sql(
        "UPDATE t "
        "SET label = CASE WHEN score > 10 THEN 'a' ELSE 'b' END "
        "WHERE CASE WHEN score > 10 THEN 'a' ELSE 'b' END = 'a'"
    )
    assert parsed["set"][0]["value"]["type"] == "case"
    where_condition = parsed["where"]["conditions"][0]
    assert where_condition["column"]["type"] == "case"


def test_searched_case_select_with_alias(tmp_path: Path) -> None:
    file_path = tmp_path / "case_searched_alias.xlsx"
    _create_case_workbook(file_path)

    rows, description = _query(
        file_path,
        "SELECT id, CASE WHEN score >= 10 THEN 'high' WHEN score >= 5 THEN 'mid' ELSE 'low' END AS band "
        "FROM t ORDER BY id",
    )
    assert description == ["id", "band"]
    assert rows == [(1, "high"), (2, "mid"), (3, "low"), (4, "high")]


def test_simple_case_select_and_implicit_alias(tmp_path: Path) -> None:
    file_path = tmp_path / "case_simple_alias.xlsx"
    _create_case_workbook(file_path)

    rows, description = _query(
        file_path,
        "SELECT CASE tier WHEN 'gold' THEN 1 WHEN 'silver' THEN 2 ELSE 3 END rank FROM t ORDER BY id",
    )
    assert description == ["rank"]
    assert rows == [(1,), (2,), (3,), (1,)]


def test_case_update_set_value(tmp_path: Path) -> None:
    file_path = tmp_path / "case_update.xlsx"
    _create_case_workbook(file_path)

    rowcount = _execute(
        file_path,
        "UPDATE t SET label = CASE WHEN score >= 10 THEN 'keep' ELSE 'drop' END",
    )
    assert rowcount == 4

    rows, _ = _query(file_path, "SELECT id, label FROM t ORDER BY id")
    assert rows == [(1, "keep"), (2, "drop"), (3, "drop"), (4, "keep")]


def test_case_null_and_no_else_behavior(tmp_path: Path) -> None:
    file_path = tmp_path / "case_null.xlsx"
    _create_case_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT id, CASE WHEN score IS NULL THEN 'missing' ELSE NULL END AS note FROM t ORDER BY id",
    )
    assert rows == [(1, None), (2, None), (3, "missing"), (4, None)]

    rows, _ = _query(
        file_path,
        "SELECT CASE WHEN score > 100 THEN 'x' END AS missing_case FROM t WHERE id = 1",
    )
    assert rows == [(None,)]


def test_case_with_arithmetic_and_nested_case(tmp_path: Path) -> None:
    file_path = tmp_path / "case_arithmetic_nested.xlsx"
    _create_case_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT id, CASE WHEN score >= 10 THEN score + points ELSE points * 2 END AS calc FROM t ORDER BY id",
    )
    assert rows == [(1, 17.0), (2, 6.0), (3, 8.0), (4, 12.0)]

    rows, _ = _query(
        file_path,
        "SELECT id, "
        "CASE WHEN score >= 10 THEN CASE WHEN points > 1 THEN 'A' ELSE 'B' END ELSE 'C' END AS nested_value "
        "FROM t ORDER BY id",
    )
    assert rows == [(1, "A"), (2, "C"), (3, "C"), (4, "B")]


def test_case_in_where_operand(tmp_path: Path) -> None:
    file_path = tmp_path / "case_where_operand.xlsx"
    _create_case_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT id FROM t "
        "WHERE CASE WHEN score > 10 THEN 'a' ELSE 'b' END = 'a' "
        "ORDER BY id",
    )
    assert rows == [(1,), (4,)]


def test_case_parameter_binding_select_and_update(tmp_path: Path) -> None:
    file_path = tmp_path / "case_params.xlsx"
    _create_case_workbook(file_path)

    rows, _ = _query(
        file_path,
        "SELECT CASE WHEN score >= ? THEN ? ELSE ? END AS bucket FROM t WHERE id = ?",
        (10, "high", "low", 2),
    )
    assert rows == [("low",)]

    rowcount = _execute(
        file_path,
        "UPDATE t SET label = CASE WHEN score >= ? THEN ? ELSE ? END WHERE id = ?",
        (10, "high", "low", 1),
    )
    assert rowcount == 1

    rows, _ = _query(file_path, "SELECT label FROM t WHERE id = 1")
    assert rows == [("high",)]
