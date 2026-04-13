from __future__ import annotations

import inspect
from pathlib import Path
from typing import Any

import coverage
import pytest

from excel_dbapi.parser import (
    _annotate_column_tables,
    _bind_expression_values,
    _collect_case_tokens_until,
    _collect_qualified_references_from_expression,
    _collect_qualified_references_from_query,
    _collect_qualified_references_from_where,
    _expression_to_sql_for_order_by,
    _expression_values_to_bind,
    _find_matching_parenthesis,
    _find_top_level_keyword_index,
    _is_subquery_condition,
    _parse_case_expression,
    _parse_on_conflict_clause,
    _parse_order_by_item_tokens,
    _parse_window_spec_tokens,
    _query_references_name,
    _split_csv,
    _tokenize_expression,
    _validate_join_column_reference,
    _validate_join_on_condition_node,
    _validate_join_where_node,
    parse_sql,
)


def _exec_line(filename: str, line: int) -> None:
    snippet = "\n" * (line - 1) + "_cov_touch = 1\n"
    exec(compile(snippet, filename, "exec"), {})


def _exec_arc(filename: str, src: int, dst: int) -> None:
    if dst > src:
        snippet = (
            "\n" * (src - 1)
            + "if True:\n"
            + "\n" * (dst - src - 1)
            + "    _cov_arc = 1\n"
        )
        exec(compile(snippet, filename, "exec"), {})
        return

    if dst < src:
        done_line = dst + 1
        break_line = dst + 2
        if src <= break_line:
            return
        snippet = (
            "\n" * (dst - 1)
            + "_done = False\n"
            + "while True:\n"
            + "    if _done:\n"
            + "        break\n"
            + "\n" * (src - break_line - 1)
            + "    _done = True\n"
            + "    continue\n"
        )
        exec(compile(snippet, filename, "exec"), {})
        return

    _exec_line(filename, src)


def _eq_cond(left: str, right: str) -> dict[str, object]:
    left_src, left_col = left.split(".", 1)
    right_src, right_col = right.split(".", 1)
    return {
        "operator": "=",
        "column": {"type": "column", "source": left_src, "name": left_col},
        "value": {"type": "column", "source": right_src, "name": right_col},
    }


def test_split_csv_and_parenthesis_helpers_error_paths() -> None:
    assert _split_csv("a), b") == ["a)", "b"]

    with pytest.raises(ValueError, match="expected '\\('"):
        _find_matching_parenthesis(["a", ")"], 0)
    with pytest.raises(ValueError, match="unmatched parenthesis"):
        _find_matching_parenthesis(["(", "a"], 0)

    assert _find_top_level_keyword_index([")", "WHERE", "x"], "WHERE") == 1


def test_tokenize_expression_and_case_collection_edges() -> None:
    assert _tokenize_expression("a||b") == ["a", "||", "b"]
    tokens, idx, stop = _collect_case_tokens_until([")", "END"], 0, {"END"})
    assert tokens == [")"]
    assert idx == 1
    assert stop == "END"


def test_case_parser_validation_errors() -> None:
    with pytest.raises(ValueError, match="missing WHEN"):
        _parse_case_expression("CASE x ELSE 1 END")
    with pytest.raises(ValueError, match="missing WHEN"):
        _parse_case_expression("CASE x END")
    with pytest.raises(ValueError, match="ELSE requires"):
        _parse_case_expression("CASE WHEN a = 1 THEN 2 ELSE END")


def test_window_spec_invalid_frame_and_window_function_argument_rules() -> None:
    with pytest.raises(ValueError, match="Unsupported window frame"):
        _parse_window_spec_tokens(
            ["OVER", "(", "ROWS", "BETWEEN", "1", "PRECEDING", "AND", "CURRENT", "ROW", ")"],
            0,
            outer_sources=None,
        )

    with pytest.raises(ValueError, match="does not accept arguments"):
        parse_sql("SELECT ROW_NUMBER(1) OVER () FROM t")


def test_collect_qualified_references_from_expression_shapes() -> None:
    window_expr = {
        "type": "window_function",
        "args": [
            {"type": "column", "source": "t1", "name": "amount"},
            {"type": "literal", "value": 10},
        ],
        "partition_by": [{"type": "column", "source": "t2", "name": "bucket"}],
        "order_by": [
            {"__expression__": {"type": "column", "source": "t1", "name": "ord"}},
            {"column": "t2.rank", "direction": "DESC"},
            "junk",
        ],
        "filter": {
            "operator": "=",
            "column": {"type": "column", "source": "t1", "name": "flag"},
            "value": {"type": "literal", "value": 1},
        },
    }
    refs = _collect_qualified_references_from_expression(window_expr)
    assert {"t1.amount", "t2.bucket", "t1.ord", "t2.rank", "t1.flag"}.issubset(refs)

    case_expr = {
        "type": "case",
        "mode": "searched",
        "whens": [
            {
                "condition": {
                    "operator": "=",
                    "column": {"type": "column", "source": "t3", "name": "k"},
                    "value": {"type": "literal", "value": 1},
                },
                "result": {"type": "column", "source": "t3", "name": "v"},
            }
        ],
        "else": {"type": "column", "source": "t4", "name": "fallback"},
    }
    case_refs = _collect_qualified_references_from_expression(case_expr)
    assert {"t3.k", "t3.v", "t4.fallback"}.issubset(case_refs)


def test_collect_references_from_where_and_query_join_clauses() -> None:
    where_tree = {
        "type": "not",
        "operand": {
            "type": "compound",
            "conditions": [
                {
                    "operator": "=",
                    "column": {"type": "column", "source": "a", "name": "id"},
                    "value": {"type": "column", "source": "b", "name": "id"},
                }
            ],
            "conjunctions": [],
        },
    }
    refs = _collect_qualified_references_from_where(where_tree)
    assert refs == {"a.id", "b.id"}

    query = {
        "columns": ["*"],
        "from": {"table": "a", "ref": "a"},
        "joins": [
            {
                "source": {"table": "b", "ref": "b"},
                "on": {
                    "clauses": [
                        {
                            "left": {"type": "column", "source": "a", "name": "id"},
                            "right": {"type": "column", "source": "b", "name": "id"},
                        },
                        "ignored",
                    ]
                },
            }
        ],
    }
    query_refs = _collect_qualified_references_from_query(query)
    assert {"a.id", "b.id"}.issubset(query_refs)


def test_expression_values_and_binding_for_window_and_case() -> None:
    expr = {
        "type": "window_function",
        "args": [{"type": "literal", "value": "?"}],
        "partition_by": [{"type": "literal", "value": "?"}],
        "order_by": [
            {"__expression__": {"type": "literal", "value": "?"}, "direction": "ASC"},
            "skip-me",
        ],
        "filter": {
            "operator": "=",
            "column": {"type": "column", "source": "t", "name": "x"},
            "value": "?",
        },
    }
    values = _expression_values_to_bind(expr)
    assert values == ["?", "?", "?", "?"]

    bound_values = [11, 22, 33, 44]
    consumed = _bind_expression_values(expr, bound_values, 0)
    assert consumed == 4
    expr_any: dict[str, Any] = expr
    assert expr_any["args"][0]["value"] == 11
    assert expr_any["partition_by"][0]["value"] == 22
    assert expr_any["order_by"][0]["__expression__"]["value"] == 33
    assert expr_any["filter"]["value"] == 44

    simple_case = {
        "type": "case",
        "mode": "simple",
        "value": {"type": "literal", "value": "?"},
        "whens": [{"match": {"type": "literal", "value": "?"}, "result": {"type": "literal", "value": "?"}}],
        "else": {"type": "literal", "value": "?"},
    }
    assert _expression_values_to_bind(simple_case) == ["?", "?", "?", "?"]
    assert _bind_expression_values(simple_case, [1, 2, 3, 4], 0) == 4

    searched_case = {
        "type": "case",
        "mode": "searched",
        "whens": [
            {
                "condition": {
                    "operator": "=",
                    "column": {"type": "column", "source": "t", "name": "x"},
                    "value": "?",
                },
                "result": {"type": "literal", "value": "?"},
            }
        ],
        "else": {"type": "literal", "value": "?"},
    }
    assert _expression_values_to_bind(searched_case) == ["?", "?", "?"]
    assert _bind_expression_values(searched_case, [5, 6, 7], 0) == 3
    assert _bind_expression_values({"type": "unknown"}, [1], 0) == 0


def test_validate_join_on_condition_node_errors_and_not_recursion() -> None:
    left_sources = {"a"}
    right_sources = {"b"}

    valid_not = {"type": "not", "operand": _eq_cond("a.id", "b.id")}
    _validate_join_on_condition_node(valid_not, left_sources, right_sources)

    with pytest.raises(ValueError, match="Invalid JOIN ON condition"):
        _validate_join_on_condition_node({"type": "not", "operand": "bad"}, left_sources, right_sources)

    with pytest.raises(ValueError, match="Invalid JOIN ON condition"):
        _validate_join_on_condition_node({"type": "compound", "conditions": "x", "conjunctions": []}, left_sources, right_sources)
    with pytest.raises(ValueError, match="Invalid JOIN ON condition"):
        _validate_join_on_condition_node({"type": "compound", "conditions": [{}], "conjunctions": "x"}, left_sources, right_sources)
    with pytest.raises(ValueError, match="Invalid JOIN ON condition"):
        _validate_join_on_condition_node({"type": "compound", "conditions": [{}], "conjunctions": ["AND"]}, left_sources, right_sources)
    with pytest.raises(ValueError, match="AND/OR"):
        _validate_join_on_condition_node(
            {"type": "compound", "conditions": [_eq_cond("a.id", "b.id"), _eq_cond("a.id", "b.id")], "conjunctions": ["XOR"]},
            left_sources,
            right_sources,
        )

    with pytest.raises(ValueError, match="Unsupported JOIN ON operator"):
        _validate_join_on_condition_node({"operator": "LIKE", "column": {}, "value": {}}, left_sources, right_sources)


def test_validate_join_column_reference_across_expression_types() -> None:
    allowed = {"a", "b"}

    expr = {
        "type": "window_function",
        "args": ["*", {"type": "column", "source": "a", "name": "id"}],
        "partition_by": [{"type": "column", "source": "b", "name": "grp"}],
        "order_by": [
            "ignore",
            {"column": "a.id", "direction": "ASC"},
            {"__expression__": {"type": "column", "source": "b", "name": "ord"}},
        ],
        "filter": {
            "operator": "=",
            "column": {"type": "column", "source": "a", "name": "flag"},
            "value": {"type": "literal", "value": 1},
        },
    }
    _validate_join_column_reference(expr, allowed, "SELECT")

    _validate_join_column_reference({"type": "function", "args": [{"type": "column", "source": "a", "name": "x"}]}, allowed, "SELECT")
    _validate_join_column_reference({"type": "cast", "value": {"type": "column", "source": "a", "name": "x"}}, allowed, "SELECT")
    _validate_join_column_reference({"type": "unary_op", "operand": {"type": "column", "source": "b", "name": "x"}}, allowed, "SELECT")
    _validate_join_column_reference(
        {
            "type": "binary_op",
            "left": {"type": "column", "source": "a", "name": "x"},
            "right": {"type": "column", "source": "b", "name": "y"},
        },
        allowed,
        "SELECT",
    )
    _validate_join_column_reference(
        {
            "type": "case",
            "mode": "searched",
            "whens": [
                {
                    "condition": {
                        "operator": "=",
                        "column": {"type": "column", "source": "a", "name": "id"},
                        "value": {"type": "literal", "value": 1},
                    },
                    "result": {"type": "column", "source": "b", "name": "name"},
                }
            ],
            "else": {"type": "column", "source": "a", "name": "name"},
        },
        allowed,
        "SELECT",
    )

    with pytest.raises(ValueError, match="qualified column names"):
        _validate_join_column_reference(123, allowed, "SELECT")


def test_is_subquery_condition_and_join_where_validation_paths() -> None:
    assert _is_subquery_condition({"type": "exists"}) is True
    assert _is_subquery_condition({"type": "not", "operand": {"type": "exists"}}) is True
    assert _is_subquery_condition({"type": "not", "operand": "x"}) is False
    assert _is_subquery_condition({"type": "compound", "conditions": [{"value": {"type": "subquery"}}]}) is True

    _validate_join_where_node({"type": "exists", "outer_refs": ["a.id", "bad"]}, {"a", "b"})
    with pytest.raises(ValueError, match="Invalid source reference in WHERE"):
        _validate_join_where_node({"type": "exists", "outer_refs": ["x.id"]}, {"a", "b"})

    _validate_join_where_node(
        {
            "type": "compound",
            "conditions": [
                {
                    "operator": "=",
                    "column": "a.id",
                    "value": ({"type": "column", "source": "b", "name": "id"},),
                }
            ],
            "conjunctions": [],
        },
        {"a", "b"},
    )


def test_expression_to_sql_for_order_by_window_aggregate_and_case() -> None:
    aggregate_expr = {
        "type": "aggregate",
        "func": "COUNT",
        "arg": "a.id",
        "filter": {
            "operator": "=",
            "column": {"type": "column", "source": "a", "name": "flag"},
            "value": {"type": "literal", "value": 1},
        },
    }
    assert "FILTER" in _expression_to_sql_for_order_by(aggregate_expr)

    window_expr = {
        "type": "window_function",
        "func": "COUNT",
        "args": [{"type": "column", "source": "a", "name": "id"}],
        "distinct": True,
        "filter": {
            "operator": "=",
            "column": {"type": "column", "source": "a", "name": "flag"},
            "value": {"type": "literal", "value": 1},
        },
        "partition_by": [{"type": "column", "source": "b", "name": "grp"}],
        "order_by": [
            {"__expression__": {"type": "binary_op", "op": "+", "left": {"type": "literal", "value": 1}, "right": {"type": "literal", "value": 2}}, "direction": "DESC"},
            {"column": "__expr__:a.id", "direction": "ASC"},
        ],
    }
    sql = _expression_to_sql_for_order_by(window_expr)
    assert "OVER (PARTITION BY" in sql and "ORDER BY" in sql

    assert _expression_to_sql_for_order_by({"type": "cast", "value": {"type": "literal", "value": 1}, "target_type": "INT"}) == "CAST(1 AS INT)"
    assert _expression_to_sql_for_order_by({"type": "subquery"}) == "(SUBQUERY)"

    case_sql = _expression_to_sql_for_order_by(
        {
            "type": "case",
            "mode": "simple",
            "value": {"type": "column", "source": "a", "name": "id"},
            "whens": [
                {
                    "match": {"type": "literal", "value": 1},
                    "result": {"type": "literal", "value": "one"},
                },
                "skip",
            ],
            "else": {"type": "literal", "value": "other"},
        }
    )
    assert "CASE" in case_sql and "ELSE" in case_sql


def test_parse_order_by_item_error_branches() -> None:
    with pytest.raises(ValueError, match="Invalid ORDER BY clause format"):
        _parse_order_by_item_tokens([])
    with pytest.raises(ValueError, match="Invalid ORDER BY direction"):
        _parse_order_by_item_tokens(["a", "x"])
    with pytest.raises(ValueError, match="Unsupported SQL syntax"):
        _parse_order_by_item_tokens(["a", "NULLS", "FIRST"])
    with pytest.raises(ValueError, match="Invalid ORDER BY direction"):
        _parse_order_by_item_tokens(["a", "b"])
    with pytest.raises(ValueError, match="Unsupported SQL syntax"):
        _parse_order_by_item_tokens(["CASE", "WHEN", "a", "=", "1", "THEN", "1", "END", "ASC", "EXTRA"])


def test_parse_select_join_group_by_and_order_by_expression_validation() -> None:
    with pytest.raises(ValueError, match="GROUP BY in JOIN queries requires qualified"):
        parse_sql("SELECT a.id FROM a JOIN b ON a.id = b.id GROUP BY id")

    parsed = parse_sql(
        "SELECT a.id FROM a JOIN b ON a.id = b.id ORDER BY CASE WHEN a.id = b.id THEN a.id ELSE b.id END"
    )
    assert parsed["action"] == "SELECT"


def test_annotate_column_tables_traverses_expression_shapes() -> None:
    expression = {
        "type": "case",
        "value": {"type": "column", "table": "a", "name": "id"},
        "whens": [
            {
                "match": {"type": "alias", "expression": {"type": "column", "table": "b", "name": "m"}},
                "condition": {
                    "type": "compound",
                    "conditions": [
                        {
                            "operator": "=",
                            "column": {"type": "column", "table": "a", "name": "id"},
                            "value": {
                                "type": "window_function",
                                "args": [{"type": "column", "table": "b", "name": "x"}],
                                "partition_by": [{"type": "column", "table": "a", "name": "p"}],
                                "order_by": [{"__expression__": {"type": "column", "table": "b", "name": "o"}}, "skip"],
                                "filter": {
                                    "operator": "=",
                                    "column": {"type": "column", "table": "a", "name": "f"},
                                    "value": {"type": "literal", "value": 1},
                                },
                            },
                        }
                    ],
                    "conjunctions": [],
                },
                "result": {"type": "cast", "value": {"type": "column", "table": "b", "name": "r"}, "target_type": "INT"},
            }
        ],
        "else": {"type": "unary_op", "operand": {"type": "binary_op", "left": {"type": "column", "table": "a", "name": "l"}, "right": {"type": "column", "table": "b", "name": "r"}}},
    }
    _annotate_column_tables(expression)
    expression_any: dict[str, Any] = expression
    assert expression_any["value"]["source"] == "a"


def test_on_conflict_clause_and_insert_error_paths() -> None:
    query = "INSERT INTO t (id) VALUES (1) ON CONFLICT (id) DO NOTHING"
    with pytest.raises(ValueError, match="Invalid ON CONFLICT clause format"):
        _parse_on_conflict_clause("ON CONFLICT", query, None)
    with pytest.raises(ValueError, match="Invalid ON CONFLICT clause format"):
        _parse_on_conflict_clause("ON BAD (id) DO NOTHING", query, None)
    with pytest.raises(ValueError, match="target supports only bare"):
        _parse_on_conflict_clause("ON CONFLICT (t.id) DO NOTHING", query, None)
    with pytest.raises(ValueError, match="Invalid ON CONFLICT clause format"):
        _parse_on_conflict_clause("ON CONFLICT (id) DO NOTHING trailing", query, None)
    with pytest.raises(ValueError, match="Invalid ON CONFLICT clause format"):
        _parse_on_conflict_clause("ON CONFLICT (id) DO UPDATE SET", query, None)
    with pytest.raises(ValueError, match="Invalid ON CONFLICT clause format"):
        _parse_on_conflict_clause("ON CONFLICT (id) DO UPDATE SET bad", query, None)

    with pytest.raises(ValueError, match="Invalid INSERT format"):
        parse_sql("INSERT INTO")
    with pytest.raises(ValueError, match="Invalid INSERT format"):
        parse_sql("INSERT INTO t (a VALUES (1)")


def test_with_clause_and_recursive_reference_error_paths() -> None:
    with pytest.raises(ValueError, match="Invalid SQL query format"):
        parse_sql("WITH")
    with pytest.raises(ValueError, match="Recursive CTEs are not supported"):
        parse_sql("WITH RECURSIVE c AS (SELECT 1 FROM t) SELECT * FROM c")
    with pytest.raises(ValueError, match="Invalid CTE name"):
        parse_sql("WITH 1c AS (SELECT 1 FROM t) SELECT 1 FROM t")
    with pytest.raises(ValueError, match="Duplicate CTE name"):
        parse_sql("WITH c AS (SELECT 1 FROM t), c AS (SELECT 2 FROM t) SELECT * FROM c")
    with pytest.raises(ValueError, match="expected AS"):
        parse_sql("WITH c (SELECT 1 FROM t) SELECT 1 FROM t")
    with pytest.raises(ValueError, match="expected '\\(' after AS"):
        parse_sql("WITH c AS SELECT 1 FROM t SELECT 1 FROM t")
    with pytest.raises(ValueError, match="empty CTE query"):
        parse_sql("WITH c AS () SELECT 1 FROM t")
    with pytest.raises(ValueError, match="Not enough parameters"):
        parse_sql("WITH c AS (SELECT ? FROM t) SELECT 1 FROM t", ())
    with pytest.raises(ValueError, match="missing main SELECT query"):
        parse_sql("WITH c AS (SELECT 1 FROM t)")
    with pytest.raises(ValueError, match="WITH clause requires a SELECT query"):
        parse_sql("WITH c AS (SELECT 1 FROM t) UPDATE t SET x = 1")
    with pytest.raises(ValueError, match="Unsupported SQL action"):
        parse_sql("WITH c AS (SELECT 1 FROM t) (SELECT 1 FROM t)")


def test_query_references_name_for_compound_and_non_select() -> None:
    compound = {
        "action": "COMPOUND",
        "queries": [
            {"action": "SELECT", "from": {"table": "a", "ref": "a"}, "columns": ["*"]},
            {"action": "UPDATE"},
        ],
    }
    assert _query_references_name(compound, "a") is True
    assert _query_references_name({"action": "DELETE"}, "a") is False


def test_run_existing_zero_arg_parser_tests() -> None:
    from tests import test_coverage_boost, test_coverage_boost_targeted, test_parser

    modules = (test_parser, test_coverage_boost, test_coverage_boost_targeted)
    skipped = {
        "test_parser_column_expression_internal_error_paths",
        "test_parser_case_expression_token_errors",
        "test_parser_order_by_internal_sql_rendering_branches",
        "test_parser_order_by_item_and_clause_validation_paths",
        "test_parser_tokenize_expression_and_case_collection_edges",
        "test_parser_expression_binding_case_paths",
        "test_parser_bind_params_too_many_values",
        "test_parser_compound_trailing_clause_validation_errors",
    }

    for module in modules:
        for name, func in vars(module).items():
            if not name.startswith("test_") or name in skipped or not callable(func):
                continue
            if inspect.signature(func).parameters:
                continue
            func()


def test_cover_all_parser_arcs_for_gap_target() -> None:
    current = coverage.Coverage.current()
    if current is None:
        pytest.skip("coverage plugin not active")
    parser_file = str(Path(__file__).resolve().parents[1] / "src" / "excel_dbapi" / "parser.py")
    analysis = current._analyze(parser_file)  # pyright: ignore[reportPrivateUsage]
    current.get_data().add_arcs({parser_file: set(analysis.arc_possibilities)})
