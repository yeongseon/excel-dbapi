from datetime import date
from pathlib import Path

import pytest
from excel_dbapi.exceptions import DatabaseError
from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection
from excel_dbapi.engines.base import TableData
from excel_dbapi.engines.openpyxl.backend import OpenpyxlBackend
from excel_dbapi.exceptions import ProgrammingError
from excel_dbapi.executor import SharedExecutor


def _create_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "t"
    sheet.append(["id", "name", "txt", "d", "v"])
    sheet.append([1, "alpha", "10", date(2024, 1, 15), 10])
    sheet.append([2, "beta", "x", date(2024, 1, 16), 20])
    sheet.append([3, None, None, None, None])

    u = workbook.create_sheet("u")
    u.append(["id", "v2"])
    u.append([1, 100])
    u.append([2, 200])

    empty = workbook.create_sheet("empty")
    empty.append(["id", "value"])

    workbook.save(path)


def test_scalar_edges_and_date_part_null_paths(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_scalar_edges.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT "
            "COALESCE(NULL, NULL), "
            "SUBSTR('hello', -2), "
            "SUBSTR('hello', 0, 2), "
            "SUBSTR('hello', 2, 0), "
            "SUBSTR('hello', 2, NULL), "
            "YEAR(NULL), MONTH(NULL), DAY(NULL), "
            "YEAR(d), MONTH('2024-01-31'), DAY('2024-01-15T11:22:33Z') "
            "FROM t WHERE id = 1"
        )
        assert cur.fetchall() == [
            (None, "lo", "he", "", None, None, None, None, 2024, 1, 15)
        ]


def test_scalar_argument_validation_paths(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_scalar_validate.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cur = conn.cursor()
        with pytest.raises(ProgrammingError, match="Invalid arguments for SUBSTR"):
            cur.execute("SELECT SUBSTR('abc', '', 1) FROM t")

        with pytest.raises(ProgrammingError, match="Unsupported function"):
            cur.execute("SELECT MADE_UP_FN(1) FROM t")

        with pytest.raises(ProgrammingError, match="expects at least"):
            conn._executor._call_function("LENGTH", [])

        with pytest.raises(ProgrammingError, match="expects at most"):
            conn._executor._call_function("NULLIF", [1, 1, 1])

        with pytest.raises(ProgrammingError, match="Invalid arguments for SUBSTR"):
            conn._executor._call_function("SUBSTR", ["abc", True, 1])


def test_like_escape_error_paths(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_like_escape.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cur = conn.cursor()
        with pytest.raises(
            ProgrammingError, match="ESCAPE requires a single character"
        ):
            cur.execute("SELECT id FROM t WHERE name LIKE 'a%' ESCAPE 'xx'")

        with pytest.raises(ProgrammingError, match="trailing ESCAPE character"):
            cur.execute("SELECT id FROM t WHERE name LIKE 'abc!' ESCAPE '!'")


def test_cast_error_and_conversion_paths(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_cast_paths.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cur = conn.cursor()

        with pytest.raises(ProgrammingError, match="Cannot cast empty string to DATE"):
            cur.execute("SELECT CAST('' AS DATE) FROM t")

        cur.execute("SELECT CAST('2024-01-15Z' AS DATE) FROM t LIMIT 1")
        assert cur.fetchone() == (date(2024, 1, 15),)

        with pytest.raises(ProgrammingError, match="Unsupported CAST target type"):
            cur.execute("SELECT CAST(1 AS BLOB) FROM t")


def test_window_rank_without_order_and_running_sum(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_window_paths.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, RANK() OVER () FROM t ORDER BY id")
        assert cur.fetchall() == [(1, 1), (2, 1), (3, 1)]

        cur.execute("SELECT id, SUM(v) OVER (ORDER BY id) FROM t ORDER BY id")
        assert cur.fetchall() == [(1, 10.0), (2, 30.0), (3, 30.0)]


def test_window_programming_error_paths(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_window_errors.xlsx"
    _create_workbook(file_path)
    backend = OpenpyxlBackend(str(file_path))
    executor = SharedExecutor(backend)
    rows = [
        {"t.id": 1, "id": 1, "v": 10},
        {"t.id": 2, "id": 2, "v": 20},
    ]

    with pytest.raises(ProgrammingError, match="requires an argument"):
        executor._apply_window_functions(
            rows,
            [{"type": "window_function", "func": "SUM", "args": [], "order_by": []}],
            None,
        )

    with pytest.raises(ProgrammingError, match="Unsupported window function"):
        executor._apply_window_functions(
            rows,
            [
                {
                    "type": "window_function",
                    "func": "MEDIAN",
                    "args": ["v"],
                    "order_by": [],
                }
            ],
            None,
        )


def test_expression_and_where_sql_helpers() -> None:
    window_expr = {
        "type": "window_function",
        "func": "SUM",
        "args": ["t.v"],
        "distinct": True,
        "filter": {"column": "t.id", "operator": ">", "value": 0},
        "order_by": [
            "bad-order-item",
            {"__expression__": {"type": "column", "source": "t", "name": "id"}},
            {"column": "__expr__:t.v + 1", "direction": "DESC"},
        ],
    }
    sql = SharedExecutor._expression_to_sql(window_expr)
    assert "DISTINCT" in sql
    assert "FILTER" in sql
    assert "OVER" in sql

    case_sql = SharedExecutor._expression_to_sql(
        {
            "type": "case",
            "mode": "searched",
            "whens": [
                {
                    "condition": {"type": "exists"},
                    "result": {"type": "literal", "value": 1},
                }
            ],
            "else": {"type": "literal", "value": 0},
        }
    )
    assert case_sql.startswith("CASE")

    assert (
        SharedExecutor._where_operand_to_sql({"type": "exists"}, is_column=False)
        == "EXISTS (SUBQUERY)"
    )
    assert SharedExecutor._where_to_sql({"type": "exists"}) == "EXISTS (SUBQUERY)"
    assert SharedExecutor._contains_arithmetic_expression({"type": "mystery"}) is False


def test_collect_expression_refs_paths() -> None:
    expr = {
        "type": "window_function",
        "args": ["t.v", "*", {"type": "column", "source": "u", "name": "v2"}],
        "partition_by": [{"type": "column", "source": "t", "name": "id"}],
        "order_by": [
            "noise",
            {"__expression__": {"type": "column", "source": "t", "name": "v"}},
            {"column": "t.id"},
        ],
        "filter": {
            "column": {"type": "column", "source": "u", "name": "id"},
            "operator": ">",
            "value": 0,
        },
    }
    refs = SharedExecutor._collect_expression_column_refs(expr)
    assert "t.v" in refs
    assert "u.v2" in refs
    assert "t.id" in refs
    assert "u.id" in refs

    aggregate_refs = SharedExecutor._collect_expression_column_refs(
        {
            "type": "aggregate",
            "arg": "t.v",
            "filter": {"type": "exists"},
        }
    )
    assert aggregate_refs == {"t.v"}


def test_collect_where_refs_paths() -> None:
    where = {
        "conditions": [
            {"type": "exists"},
            {
                "column": {"type": "column", "source": "t", "name": "id"},
                "operator": "IN",
                "value": [
                    {"type": "column", "source": "u", "name": "id"},
                    1,
                ],
            },
        ]
    }
    refs = SharedExecutor._collect_where_column_refs(where)
    assert "t.id" in refs
    assert "u.id" in refs


def test_aggregate_select_validation_paths(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_agg_validation.xlsx"
    _create_workbook(file_path)
    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cur = conn.cursor()

        with pytest.raises(ProgrammingError, match="Non-aggregate columns"):
            cur.execute("SELECT id, COUNT(*) FROM t")

        with pytest.raises(ProgrammingError, match=r"Unknown column\(s\): nope"):
            cur.execute("SELECT SUM(v) FROM t GROUP BY nope")

        with pytest.raises(ProgrammingError, match="Unknown column: t.nope"):
            cur.execute("SELECT SUM(t.nope) FROM t")

        with pytest.raises(ProgrammingError, match="must be a GROUP BY column"):
            cur.execute("SELECT id, COUNT(*) FROM t GROUP BY id HAVING name = 'alpha'")


def test_aggregate_distinct_dedupe_path(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_agg_distinct.xlsx"
    _create_workbook(file_path)
    backend = OpenpyxlBackend(str(file_path))
    executor = SharedExecutor(backend)
    data = backend.read_sheet("t")
    headers = list(data.headers)
    rows = [
        executor._build_scoped_row(executor._row_from_values(headers, list(r)))
        for r in data.rows
    ]
    result = executor._execute_aggregate_select(
        "SELECT",
        {"distinct": True, "order_by": None, "limit": None, "offset": None},
        headers,
        rows,
        [{"type": "aggregate", "func": "COUNT", "arg": "*"}],
        None,
        None,
    )
    assert result.rows == [(3,)]


def test_join_validation_paths_with_custom_ast(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_join_validation.xlsx"
    _create_workbook(file_path)
    backend = OpenpyxlBackend(str(file_path))
    executor = SharedExecutor(backend)
    left_data = backend.read_sheet("t")

    parsed = {
        "action": "SELECT",
        "table": "t",
        "from": {"table": "t", "ref": "t"},
        "joins": [
            {
                "type": "INNER",
                "source": {"table": "u", "ref": "u"},
                "on": {"column": "t.id", "operator": "=", "value": "u.id"},
            }
        ],
        "columns": [
            None,
            {"type": "literal", "value": 1},
            {
                "type": "unary_op",
                "operand": {"type": "column", "source": "t", "name": "v"},
            },
            {
                "type": "binary_op",
                "left": {"type": "column", "source": "t", "name": "v"},
                "right": {"type": "column", "source": "u", "name": "v2"},
            },
            {
                "type": "function",
                "args": [{"type": "column", "source": "t", "name": "id"}],
            },
            {"type": "cast", "value": {"type": "column", "source": "u", "name": "v2"}},
            {
                "type": "case",
                "mode": "simple",
                "value": {"type": "column", "source": "t", "name": "id"},
                "whens": [
                    "skip",
                    {
                        "match": {"type": "column", "source": "u", "name": "id"},
                        "result": {"type": "literal", "value": 1},
                    },
                ],
                "else": {"type": "subquery"},
            },
            {
                "type": "window_function",
                "args": ["*", "t.id", {"type": "column", "source": "u", "name": "v2"}],
                "partition_by": [{"type": "column", "source": "t", "name": "id"}],
                "order_by": [
                    "junk",
                    {"__expression__": {"type": "column", "source": "t", "name": "id"}},
                    {"column": "__expr__:t.id + 1"},
                    {"column": "t.id"},
                ],
                "filter": {"column": "u.id", "operator": ">", "value": 0},
            },
        ],
        "where": None,
        "group_by": None,
        "having": None,
        "order_by": [{"column": "t.id", "direction": "ASC"}],
    }

    with pytest.raises(Exception):
        executor._execute_join_select("SELECT", parsed, "t", left_data)


def test_misc_internal_paths() -> None:
    backend = OpenpyxlBackend("tests/data/sample.xlsx")
    executor = SharedExecutor(backend)

    assert (
        SharedExecutor._normalize_single_source_aggregate_arg("t.value", ["a.b"])
        == "t.value"
    )
    assert (
        SharedExecutor._normalize_single_source_aggregate_arg("t.value", ["value"])
        == "value"
    )
    assert (
        SharedExecutor._normalize_single_source_aggregate_arg("t.value", ["x"])
        == "t.value"
    )

    assert SharedExecutor._output_name({"type": "case"}) == "case_expr"
    assert SharedExecutor._source_key(
        {"type": "binary_op", "left": 1, "right": 2, "op": "+"}
    ).startswith("__expr__:")

    assert executor._sort_key(object())[0] == 0
    assert SharedExecutor._coerce_temporal_value(date(2024, 1, 1)) is not None
    assert SharedExecutor._parse_datetime_string("") is None
    assert SharedExecutor._parse_datetime_string("2024-01-01Z") is not None
    assert SharedExecutor._parse_datetime_string("2024-01-01") is not None

    row = executor._build_scoped_row({"id": 1}, headers=["id", "v"], source_refs={"t"})
    assert row["t.id"] == 1
    assert row["t.v"] is None

    with pytest.raises(DatabaseError, match="Unknown source reference"):
        executor._resolve_join_column({"t": {"id": 1}}, {"source": "x", "name": "id"})

    assert (
        executor._matches_join_on_condition({"t": {"id": 1}}, {"u": {"id": 1}}, None)
        is True
    )


def test_missing_sheet_errors_include_available_names(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_missing_sheet.xlsx"
    _create_workbook(file_path)

    with ExcelConnection(str(file_path), engine="openpyxl") as conn:
        cur = conn.cursor()
        with pytest.raises(ProgrammingError, match="Available sheets"):
            cur.execute("SELECT * FROM nope")
        with pytest.raises(ProgrammingError, match="Available sheets"):
            cur.execute("UPDATE nope SET id = 1")
        with pytest.raises(ProgrammingError, match="Available sheets"):
            cur.execute("DELETE FROM nope")
        with pytest.raises(ProgrammingError, match="Available sheets"):
            cur.execute("INSERT INTO nope VALUES (1)")


def test_insert_and_alter_validation_paths(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_insert_alter.xlsx"
    _create_workbook(file_path)
    backend = OpenpyxlBackend(str(file_path))
    executor = SharedExecutor(backend)

    parsed_insert = {
        "action": "INSERT",
        "table": "t",
        "values": 123,
    }
    with pytest.raises(DatabaseError, match="Invalid INSERT values format"):
        executor.execute(parsed_insert)

    parsed_insert_select = {
        "action": "INSERT",
        "table": "t",
        "columns": ["id"],
        "values": {
            "type": "subquery",
            "query": {
                "action": "SELECT",
                "table": "t",
                "from": {"table": "t", "ref": "t"},
                "columns": ["id", "v"],
                "where": None,
                "order_by": None,
                "limit": None,
                "offset": None,
                "distinct": False,
            },
        },
    }
    with pytest.raises(DatabaseError, match="column count mismatch"):
        executor.execute(parsed_insert_select)

    parsed_on_conflict_bad_action = {
        "action": "INSERT",
        "table": "t",
        "columns": ["id", "name", "txt", "d", "v"],
        "values": [[1, "x", "1", None, 10]],
        "on_conflict": {
            "target_columns": ["id"],
            "action": "merge",
            "set": [],
        },
    }
    with pytest.raises(DatabaseError, match="Invalid ON CONFLICT action"):
        executor.execute(parsed_on_conflict_bad_action)

    parsed_alter_bad_op = {
        "action": "ALTER",
        "table": "t",
        "operation": "NOPE",
    }
    with pytest.raises(DatabaseError, match="Unsupported ALTER operation"):
        executor.execute(parsed_alter_bad_op)


def test_direct_helper_branches_and_subquery_errors(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_helper_branches.xlsx"
    _create_workbook(file_path)
    backend = OpenpyxlBackend(str(file_path))
    executor = SharedExecutor(backend)

    with pytest.raises(ProgrammingError, match="expected numeric value"):
        executor._call_function("SUBSTR", ["abc", " ", 1])

    assert executor._call_function("SUBSTR", [None, 1]) is None
    assert executor._call_function("YEAR", [date(2024, 1, 1)]) == 2024
    with pytest.raises(ProgrammingError, match="Invalid arguments for YEAR"):
        executor._call_function("YEAR", [""])
    with pytest.raises(ProgrammingError, match="Invalid arguments for DAY"):
        executor._call_function("DAY", [object()])

    with pytest.raises(DatabaseError, match="Invalid CTE definition"):
        executor.execute(
            {"ctes": [{"name": 1, "query": {}}], "action": "SELECT", "table": "t"}
        )

    with pytest.raises(ProgrammingError, match="Unsupported subquery node type"):
        executor._eval_subquery({"type": "weird"})
    with pytest.raises(ProgrammingError, match="missing parsed query"):
        executor._eval_subquery({"type": "subquery", "mode": "scalar", "query": None})
    with pytest.raises(ProgrammingError, match="Unsupported subquery mode"):
        executor._eval_subquery(
            {
                "type": "subquery",
                "mode": "bad",
                "query": {
                    "action": "SELECT",
                    "table": "t",
                    "from": {"table": "t", "ref": "t"},
                    "columns": ["id"],
                    "where": None,
                    "order_by": None,
                    "limit": None,
                    "offset": None,
                    "distinct": False,
                },
            }
        )


def test_condition_and_expression_normalization_branches(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_condition_branches.xlsx"
    _create_workbook(file_path)
    backend = OpenpyxlBackend(str(file_path))
    executor = SharedExecutor(backend)

    row = {"x": 2, "vals": [1, 2, 3], "other": (4, 5), "z": "alpha"}
    assert (
        executor._evaluate_condition(
            row, {"column": "x", "operator": "IN", "value": None}
        )
        is False
    )
    assert (
        executor._evaluate_condition(
            row, {"column": "x", "operator": "IN", "value": [1, 2, 3]}
        )
        is True
    )
    assert (
        executor._evaluate_condition(
            row, {"column": "x", "operator": "NOT IN", "value": None}
        )
        is True
    )
    assert (
        executor._evaluate_condition(
            row, {"column": "x", "operator": "NOT IN", "value": (2, 9)}
        )
        is False
    )

    with pytest.raises(DatabaseError, match="single character"):
        executor._evaluate_condition(
            row,
            {"column": "z", "operator": "LIKE", "value": "a%", "escape": "!!"},
        )

    assert executor._eval_cast(date(2024, 1, 1), "TEXT") == "2024-01-01"
    assert executor._eval_cast("2024-01-02", "DATE") == date(2024, 1, 2)
    with pytest.raises(ProgrammingError, match="Cannot cast value"):
        executor._eval_cast(123, "DATE")

    assert executor._eval_expression(123, {}, lambda c: c) == 123
    assert (
        executor._eval_expression(
            {"type": "alias", "expression": {"type": "literal", "value": 7}},
            {},
            lambda c: c,
        )
        == 7
    )
    with pytest.raises(ProgrammingError, match="missing source"):
        executor._eval_expression({"type": "column", "name": "id"}, {}, lambda c: c)
    with pytest.raises(ProgrammingError, match="requires numeric operands"):
        executor._eval_expression(
            {
                "type": "unary_op",
                "op": "-",
                "operand": {"type": "literal", "value": "x"},
            },
            {},
            lambda c: c,
        )


def test_join_empty_headers_and_resolve_join_column_success(tmp_path: Path) -> None:
    file_path = tmp_path / "cov_join_empty_headers.xlsx"
    workbook = Workbook()
    left = workbook.active
    assert left is not None
    left.title = "left"
    workbook.create_sheet("right")
    workbook.save(file_path)

    backend = OpenpyxlBackend(str(file_path))
    executor = SharedExecutor(backend)
    left_data = TableData(headers=[], rows=[])
    parsed = {
        "from": {"table": "left", "ref": "left"},
        "joins": [],
        "columns": ["*"],
        "where": None,
        "order_by": None,
    }
    result = executor._execute_join_select("SELECT", parsed, "left", left_data)
    assert result.rows == []

    assert (
        executor._resolve_join_column({"l": {"id": 1}}, {"source": "l", "name": "id"})
        == 1
    )
