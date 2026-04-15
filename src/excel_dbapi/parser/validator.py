from __future__ import annotations

from typing import Any, Dict

from ..exceptions import SqlParseError, SqlSemanticError
from .expressions import _collect_qualified_references_from_expression
from .tokenizer import (
    _is_qualified_identifier_or_quoted,
    _parse_column_identifier,
    _split_qualified_identifier,
)


def _query_source_references(query: Dict[str, Any]) -> set[str]:
    refs: set[str] = set()
    from_entry = query.get("from")
    if isinstance(from_entry, dict):
        table_name = from_entry.get("table")
        if isinstance(table_name, str):
            refs.add(table_name)
        ref_name = from_entry.get("ref")
        if isinstance(ref_name, str):
            refs.add(ref_name)

    joins = query.get("joins")
    if isinstance(joins, list):
        for join in joins:
            if not isinstance(join, dict):
                continue
            source = join.get("source")
            if not isinstance(source, dict):
                continue
            join_table = source.get("table")
            if isinstance(join_table, str):
                refs.add(join_table)
            join_ref = source.get("ref")
            if isinstance(join_ref, str):
                refs.add(join_ref)

    return refs


def _collect_qualified_references_from_query(query: Dict[str, Any]) -> set[str]:
    from .where import _collect_qualified_references_from_where

    refs: set[str] = set()

    columns = query.get("columns")
    if isinstance(columns, list):
        for column in columns:
            refs.update(_collect_qualified_references_from_expression(column))

    where = query.get("where")
    if isinstance(where, dict):
        refs.update(_collect_qualified_references_from_where(where))

    having = query.get("having")
    if isinstance(having, dict):
        refs.update(_collect_qualified_references_from_where(having))

    group_by = query.get("group_by")
    if isinstance(group_by, list):
        for group_column in group_by:
            if isinstance(group_column, str):
                if _is_qualified_identifier_or_quoted(group_column):
                    parts = _split_qualified_identifier(group_column)
                    if parts is not None:
                        refs.add(
                            f"{_parse_column_identifier(parts[0])}.{_parse_column_identifier(parts[1])}"
                        )
                    else:
                        refs.add(group_column)
                continue
            refs.update(_collect_qualified_references_from_expression(group_column))

    order_by = query.get("order_by")
    if isinstance(order_by, list):
        for item in order_by:
            if not isinstance(item, dict):
                continue
            expression = item.get("__expression__")
            if expression is not None:
                refs.update(_collect_qualified_references_from_expression(expression))
                continue
            order_column = item.get("column")
            if isinstance(order_column, str) and _is_qualified_identifier_or_quoted(
                order_column
            ):
                parts = _split_qualified_identifier(order_column)
                if parts is not None:
                    refs.add(
                        f"{_parse_column_identifier(parts[0])}.{_parse_column_identifier(parts[1])}"
                    )
                else:
                    refs.add(order_column)

    joins = query.get("joins")
    if isinstance(joins, list):
        for join in joins:
            if not isinstance(join, dict):
                continue
            on_clause = join.get("on")
            if not isinstance(on_clause, dict):
                continue
            clauses = on_clause.get("clauses")
            if not isinstance(clauses, list):
                continue
            for clause in clauses:
                if not isinstance(clause, dict):
                    continue
                for side in ("left", "right"):
                    column = clause.get(side)
                    if not isinstance(column, dict):
                        continue
                    source = column.get("source")
                    name = column.get("name")
                    if isinstance(source, str) and isinstance(name, str):
                        refs.add(f"{source}.{name}")

    return refs


def _detect_subquery_correlation(
    query: Dict[str, Any],
    outer_sources: set[str] | None,
) -> tuple[bool, list[str]]:
    if not outer_sources:
        return False, []

    inner_sources = _query_source_references(query)
    outer_refs: set[str] = set()
    for reference in _collect_qualified_references_from_query(query):
        parts = _split_qualified_identifier(reference)
        source = (
            _parse_column_identifier(parts[0])
            if parts is not None
            else reference.split(".", 1)[0]
        )
        if source in outer_sources and source not in inner_sources:
            outer_refs.add(reference)

    ordered_outer_refs = sorted(outer_refs)
    return bool(ordered_outer_refs), ordered_outer_refs


def _parse_qualified_column_reference(token: str) -> Dict[str, str]:
    if not _is_qualified_identifier_or_quoted(token):
        raise SqlParseError(
            f"Invalid column reference in JOIN clause: {token}. Expected source.column"
        )
    parts = _split_qualified_identifier(token)
    if parts is not None:
        source = _parse_column_identifier(parts[0])
        name = _parse_column_identifier(parts[1])
    else:
        source, name = token.split(".", 1)
    return {"type": "column", "source": source, "name": name}


def _validate_join_on_condition_node(
    node: Dict[str, Any],
    left_sources: set[str],
    right_sources: set[str],
) -> None:
    if node.get("type") == "not":
        operand = node.get("operand")
        if isinstance(operand, dict):
            _validate_join_on_condition_node(operand, left_sources, right_sources)
            return
        raise SqlParseError("Invalid JOIN ON condition")

    if "conditions" in node and node.get("type") != "not":
        conditions = node.get("conditions")
        conjunctions = node.get("conjunctions")
        if not isinstance(conditions, list):
            raise SqlParseError("Invalid JOIN ON condition")
        if not isinstance(conjunctions, list):
            raise SqlParseError("Invalid JOIN ON condition")
        if len(conditions) != len(conjunctions) + 1:
            raise SqlParseError("Invalid JOIN ON condition")

        for conjunction in conjunctions:
            if str(conjunction).upper() not in {"AND", "OR"}:
                raise SqlParseError("JOIN ON supports only AND/OR conjunctions")

        for condition in conditions:
            if not isinstance(condition, dict):
                raise SqlParseError("Invalid JOIN ON condition")
            _validate_join_on_condition_node(condition, left_sources, right_sources)
        return

    operator = str(node.get("operator", "")).upper()
    if operator not in {"=", "==", "!=", "<>", ">", "<", ">=", "<="}:
        raise SqlParseError(f"Unsupported JOIN ON operator: {operator}")

    left = node.get("column")
    right = node.get("value")
    if not isinstance(left, dict) or left.get("type") != "column":
        raise SqlParseError(
            "Invalid column reference in JOIN: ON supports only qualified column-to-column comparisons"
        )
    if not isinstance(right, dict) or right.get("type") != "column":
        raise SqlParseError(
            "Invalid column reference in JOIN: ON supports only qualified column-to-column comparisons"
        )

    left_source = str(left.get("source", ""))
    right_source = str(right.get("source", ""))
    left_on_left = left_source in left_sources
    left_on_right = left_source in right_sources
    right_on_left = right_source in left_sources
    right_on_right = right_source in right_sources

    if not ((left_on_left and right_on_right) or (left_on_right and right_on_left)):
        raise SqlParseError(
            "JOIN ON references must compare columns from the two joined sources"
        )


def _validate_join_column_reference(
    column: Any,
    allowed_sources: set[str],
    context: str,
) -> None:
    if isinstance(column, str):
        if not _is_qualified_identifier_or_quoted(column):
            raise SqlSemanticError(
                f"{context} requires qualified column names in JOIN queries"
            )
        parts = _split_qualified_identifier(column)
        if parts is not None:
            source = _parse_column_identifier(parts[0])
            name = _parse_column_identifier(parts[1])
        else:
            source, name = column.split(".", 1)
        if source not in allowed_sources or not name:
            raise SqlParseError(
                f"Invalid source reference in {context}: {source}.{name}"
            )
        return

    if isinstance(column, dict):
        column_type = column.get("type")
        if column_type == "alias":
            _validate_join_column_reference(
                column.get("expression"), allowed_sources, context
            )
            return
        if column_type == "column":
            source = str(column.get("source", ""))
            name = str(column.get("name", ""))
            if source not in allowed_sources or not name:
                raise SqlParseError(
                    f"Invalid source reference in {context}: {source}.{name}"
                )
            return
        if column_type == "literal":
            return
        if column_type == "aggregate":
            arg = str(column.get("arg", ""))
            if arg != "*":
                if not _is_qualified_identifier_or_quoted(arg):
                    raise SqlParseError(
                        "Aggregate arguments in JOIN queries must be qualified column names or *"
                    )
                parts = _split_qualified_identifier(arg)
                if parts is not None:
                    source = _parse_column_identifier(parts[0])
                    name = _parse_column_identifier(parts[1])
                else:
                    source, name = arg.split(".", 1)
                if source not in allowed_sources or not name:
                    raise SqlParseError(
                        f"Invalid source reference in {context}: {source}.{name}"
                    )
            filter_clause = column.get("filter")
            if isinstance(filter_clause, dict):
                _validate_join_where_node(filter_clause, allowed_sources)
            return
        if column_type == "window_function":
            args = column.get("args")
            if isinstance(args, list):
                for argument in args:
                    if isinstance(argument, str) and argument == "*":
                        continue
                    _validate_join_column_reference(argument, allowed_sources, context)

            partition_by = column.get("partition_by")
            if isinstance(partition_by, list):
                for partition_expression in partition_by:
                    _validate_join_column_reference(
                        partition_expression, allowed_sources, context
                    )

            order_by = column.get("order_by")
            if isinstance(order_by, list):
                for item in order_by:
                    if not isinstance(item, dict):
                        continue
                    order_expression = item.get("__expression__")
                    if order_expression is not None:
                        _validate_join_column_reference(
                            order_expression, allowed_sources, context
                        )
                        continue

                    order_column = item.get("column")
                    if isinstance(order_column, str):
                        _validate_join_column_reference(
                            order_column, allowed_sources, context
                        )

            filter_clause = column.get("filter")
            if isinstance(filter_clause, dict):
                _validate_join_where_node(filter_clause, allowed_sources)
            return
        if column_type == "unary_op":
            _validate_join_column_reference(
                column.get("operand"), allowed_sources, context
            )
            return
        if column_type == "binary_op":
            _validate_join_column_reference(column.get("left"), allowed_sources, context)
            _validate_join_column_reference(column.get("right"), allowed_sources, context)
            return
        if column_type == "function":
            args = column.get("args")
            if isinstance(args, list):
                for argument in args:
                    _validate_join_column_reference(argument, allowed_sources, context)
            return
        if column_type == "cast":
            _validate_join_column_reference(column.get("value"), allowed_sources, context)
            return
        if column_type == "case":
            mode = str(column.get("mode", ""))
            if mode == "simple":
                _validate_join_column_reference(
                    column.get("value"), allowed_sources, context
                )

            whens = column.get("whens")
            if isinstance(whens, list):
                for when in whens:
                    if not isinstance(when, dict):
                        continue
                    if mode == "searched":
                        condition = when.get("condition")
                        if isinstance(condition, dict):
                            _validate_join_where_node(condition, allowed_sources)
                    else:
                        _validate_join_column_reference(
                            when.get("match"), allowed_sources, context
                        )
                    _validate_join_column_reference(
                        when.get("result"), allowed_sources, context
                    )

            else_expression = column.get("else")
            if else_expression is not None:
                _validate_join_column_reference(
                    else_expression, allowed_sources, context
                )
            return

    raise SqlSemanticError(f"{context} requires qualified column names in JOIN queries")


def _validate_join_where_columns(where: Dict[str, Any], join_sources: set[str]) -> None:
    for condition in where.get("conditions", []):
        _validate_join_where_node(condition, join_sources)


def _validate_join_where_node(node: Dict[str, Any], join_sources: set[str]) -> None:
    if node.get("type") == "exists":
        outer_refs = node.get("outer_refs")
        if isinstance(outer_refs, list):
            for reference in outer_refs:
                if not isinstance(reference, str) or "." not in reference:
                    continue
                parts = _split_qualified_identifier(reference)
                if parts is not None:
                    source = _parse_column_identifier(parts[0])
                else:
                    source = reference.split(".", 1)[0]
                if source not in join_sources:
                    raise SqlParseError(
                        f"Invalid source reference in WHERE: {reference}"
                    )
        return

    if node.get("type") == "not":
        operand = node.get("operand")
        if isinstance(operand, dict):
            _validate_join_where_node(operand, join_sources)
        return
    if "conditions" in node and node.get("type") != "not":
        for child in node["conditions"]:
            _validate_join_where_node(child, join_sources)
        return
    column_ref = node.get("column")
    if isinstance(column_ref, dict):
        _validate_join_column_reference(column_ref, join_sources, "WHERE")
    else:
        column_name = str(column_ref or "")
        if column_name:
            parsed_column = _parse_qualified_column_reference(column_name)
            _validate_join_column_reference(parsed_column, join_sources, "WHERE")

    value = node.get("value")
    if isinstance(value, dict) and value.get("type") != "subquery":
        _validate_join_column_reference(value, join_sources, "WHERE")
    elif isinstance(value, tuple):
        for part in value:
            if isinstance(part, dict):
                _validate_join_column_reference(part, join_sources, "WHERE")


def _validate_in_subquery_restrictions(subquery_parsed: Dict[str, Any]) -> None:
    if subquery_parsed.get("having") is not None:
        raise SqlParseError("HAVING is not supported in subqueries")
    if subquery_parsed.get("order_by") is not None:
        raise SqlParseError("ORDER BY is not supported in subqueries")
    if subquery_parsed.get("offset") is not None:
        raise SqlParseError("OFFSET is not supported in subqueries")
    if subquery_parsed.get("group_by") is not None:
        raise SqlParseError("GROUP BY is not supported in subqueries")
    if subquery_parsed.get("limit") is not None:
        raise SqlParseError("LIMIT is not supported in subqueries")
    if subquery_parsed.get("joins"):
        raise SqlParseError("JOIN is not supported in subqueries")
