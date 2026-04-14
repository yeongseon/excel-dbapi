from __future__ import annotations

from typing import Any, Dict, List, Optional

from ._constants import (
    _OrderByClause,
    _is_placeholder,
)
from .tokenizer import (
    _collapse_aggregate_tokens,
    _find_top_level_keyword_index,
    _is_identifier_or_quoted,
    _is_quoted_token,
    _is_qualified_identifier_or_quoted,
    _parse_column_identifier,
    _parse_table_identifier,
    _parse_value,
    _split_qualified_identifier,
    _split_csv,
    _tokenize,
)
from .expressions import (
    _aggregate_expression_to_label,
    _bind_expression_values,
    _bind_where_conditions,
    _collect_qualified_references_from_expression,
    _expression_to_sql_for_order_by,
    _expression_values_to_bind,
    _is_case_keyword,
    _normalize_aggregate_expressions,
    _parse_case_expression_tokens,
    _parse_column_expression,
    _where_values_to_bind,
)
from .where import (
    _collect_qualified_references_from_where,
    _parse_where_expression,
)


def _find_clause_positions(tokens: List[str]) -> Dict[str, int]:
    positions: Dict[str, int] = {}
    index = 0
    paren_depth = 0
    while index < len(tokens):
        token = tokens[index]
        if _is_quoted_token(token):
            index += 1
            continue

        if token == "(":
            paren_depth += 1
            index += 1
            continue
        if token == ")":
            if paren_depth > 0:
                paren_depth -= 1
            index += 1
            continue

        if paren_depth > 0:
            index += 1
            continue

        upper = token.upper()
        if upper == "WHERE" and "WHERE" not in positions:
            positions["WHERE"] = index
        elif (
            upper == "GROUP"
            and index + 1 < len(tokens)
            and tokens[index + 1].upper() == "BY"
            and "GROUP BY" not in positions
        ):
            positions["GROUP BY"] = index
            index += 1
        elif upper == "HAVING" and "HAVING" not in positions:
            positions["HAVING"] = index
        elif (
            upper == "ORDER"
            and index + 1 < len(tokens)
            and tokens[index + 1].upper() == "BY"
            and "ORDER BY" not in positions
        ):
            positions["ORDER BY"] = index
            index += 1
        elif upper == "LIMIT" and "LIMIT" not in positions:
            positions["LIMIT"] = index
        elif upper == "OFFSET" and "OFFSET" not in positions:
            positions["OFFSET"] = index

        index += 1

    return positions


def _validate_non_negative_pagination(value: Any, clause_name: str) -> None:
    if isinstance(value, int) and value < 0:
        raise ValueError(f"{clause_name} must be a non-negative integer")


def _parse_columns(
    columns_token: str,
    *,
    allow_subqueries: bool = False,
    outer_sources: set[str] | None = None,
) -> List[Any]:
    columns_token = columns_token.strip()
    if columns_token == "*":
        return ["*"]

    def _is_valid_alias_name(alias_name: str) -> bool:
        if not _is_identifier_or_quoted(alias_name):
            return False
        if _is_reserved_select_token(alias_name):
            return False
        if alias_name.upper() in {"AS", "ASC", "DESC"}:
            return False
        return True

    columns: List[Any] = []
    for raw_column in _split_csv(columns_token):
        column = raw_column.strip()
        if not column:
            continue
        collapsed_tokens = _collapse_aggregate_tokens(_tokenize(column))
        if not collapsed_tokens:
            continue

        expression_tokens = list(collapsed_tokens)
        alias_name: str | None = None
        parsed_expression: Any | None = None

        if len(collapsed_tokens) >= 3 and collapsed_tokens[-2].upper() == "AS":
            candidate_alias = collapsed_tokens[-1]
            if not _is_valid_alias_name(candidate_alias):
                raise ValueError(f"Invalid column alias: {candidate_alias}")
            expression_tokens = collapsed_tokens[:-2]
            if not expression_tokens:
                raise ValueError("Invalid column list")
            alias_name = _parse_column_identifier(candidate_alias)
        elif len(collapsed_tokens) >= 2:
            candidate_alias = collapsed_tokens[-1]
            if _is_valid_alias_name(candidate_alias):
                candidate_expression_tokens = collapsed_tokens[:-1]
                candidate_expression = " ".join(candidate_expression_tokens).strip()
                try:
                    parsed_expression = _parse_column_expression(
                        candidate_expression,
                        allow_subqueries=allow_subqueries,
                        outer_sources=outer_sources,
                    )
                except ValueError:
                    parsed_expression = None
                else:
                    expression_tokens = candidate_expression_tokens
                    alias_name = _parse_column_identifier(candidate_alias)

        if parsed_expression is None:
            expression = " ".join(expression_tokens).strip()
            parsed_expression = _parse_column_expression(
                expression,
                allow_subqueries=allow_subqueries,
                outer_sources=outer_sources,
            )

        if alias_name is None:
            columns.append(parsed_expression)
            continue

        if parsed_expression == "*":
            raise ValueError("Cannot alias wildcard column '*'")

        columns.append(
            {
                "type": "alias",
                "alias": alias_name,
                "expression": parsed_expression,
            }
        )

    if not columns:
        raise ValueError("Invalid column list")
    return columns


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
        source = _parse_column_identifier(parts[0]) if parts is not None else reference.split(".", 1)[0]
        if source in outer_sources and source not in inner_sources:
            outer_refs.add(reference)

    ordered_outer_refs = sorted(outer_refs)
    return bool(ordered_outer_refs), ordered_outer_refs


def _parse_scalar_subquery_node(
    subquery_sql: str,
    *,
    outer_sources: set[str] | None,
) -> Dict[str, Any]:
    if "?" in _tokenize(subquery_sql):
        raise ValueError(
            "Parameterized subqueries are not supported; use literal values"
        )

    parsed_subquery = _parse_select(
        subquery_sql,
        params=None,
        _allow_subqueries=True,
    )
    subquery_columns = parsed_subquery.get("columns")
    if (
        subquery_columns == ["*"]
        or not isinstance(subquery_columns, list)
        or len(subquery_columns) != 1
    ):
        raise ValueError("Scalar subquery must select exactly one column")

    correlated, outer_refs = _detect_subquery_correlation(
        parsed_subquery,
        outer_sources,
    )
    return {
        "type": "subquery",
        "query": parsed_subquery,
        "mode": "scalar",
        "correlated": correlated,
        "outer_refs": outer_refs,
    }


def _is_select_clause_token(token: str) -> bool:
    return token.upper() in {
        "WHERE",
        "GROUP",
        "HAVING",
        "ORDER",
        "LIMIT",
        "OFFSET",
        "JOIN",
        "INNER",
        "LEFT",
        "RIGHT",
        "FULL",
        "CROSS",
    }


def _is_reserved_select_token(token: str) -> bool:
    return token.upper() in {
        "WHERE",
        "GROUP",
        "HAVING",
        "ORDER",
        "LIMIT",
        "OFFSET",
        "ON",
        "JOIN",
        "INNER",
        "LEFT",
        "OUTER",
        "RIGHT",
        "FULL",
        "CROSS",
    }


def _parse_qualified_column_reference(token: str) -> Dict[str, str]:
    if not _is_qualified_identifier_or_quoted(token):
        raise ValueError(
            f"Invalid column reference in JOIN clause: {token}. Expected source.column"
        )
    parts = _split_qualified_identifier(token)
    if parts is not None:
        source = _parse_column_identifier(parts[0])
        name = _parse_column_identifier(parts[1])
    else:
        source, name = token.split(".", 1)
    return {"type": "column", "source": source, "name": name}


def _parse_join_on_condition(
    on_tokens: List[str],
    left_sources: set[str],
    right_sources: set[str],
) -> Dict[str, Any]:
    if not on_tokens:
        raise ValueError("JOIN requires ON condition")

    parsed_condition = _parse_where_expression(
        " ".join(on_tokens),
        params=None,
        bind_params=False,
        allow_aggregates=False,
        allow_subqueries=False,
    )

    _validate_join_on_condition_node(parsed_condition, left_sources, right_sources)
    return parsed_condition


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
        raise ValueError("Invalid JOIN ON condition")

    if "conditions" in node and node.get("type") != "not":
        conditions = node.get("conditions")
        conjunctions = node.get("conjunctions")
        if not isinstance(conditions, list):
            raise ValueError("Invalid JOIN ON condition")
        if not isinstance(conjunctions, list):
            raise ValueError("Invalid JOIN ON condition")
        if len(conditions) != len(conjunctions) + 1:
            raise ValueError("Invalid JOIN ON condition")

        for conjunction in conjunctions:
            if str(conjunction).upper() not in {"AND", "OR"}:
                raise ValueError("JOIN ON supports only AND/OR conjunctions")

        for condition in conditions:
            if not isinstance(condition, dict):
                raise ValueError("Invalid JOIN ON condition")
            _validate_join_on_condition_node(condition, left_sources, right_sources)
        return

    operator = str(node.get("operator", "")).upper()
    if operator not in {"=", "==", "!=", "<>", ">", "<", ">=", "<="}:
        raise ValueError(f"Unsupported JOIN ON operator: {operator}")

    left = node.get("column")
    right = node.get("value")
    if not isinstance(left, dict) or left.get("type") != "column":
        raise ValueError(
            "Invalid column reference in JOIN: ON supports only qualified column-to-column comparisons"
        )
    if not isinstance(right, dict) or right.get("type") != "column":
        raise ValueError(
            "Invalid column reference in JOIN: ON supports only qualified column-to-column comparisons"
        )

    left_source = str(left.get("source", ""))
    right_source = str(right.get("source", ""))
    left_on_left = left_source in left_sources
    left_on_right = left_source in right_sources
    right_on_left = right_source in left_sources
    right_on_right = right_source in right_sources

    if not ((left_on_left and right_on_right) or (left_on_right and right_on_left)):
        raise ValueError(
            "JOIN ON references must compare columns from the two joined sources"
        )


def _validate_join_column_reference(
    column: Any,
    allowed_sources: set[str],
    context: str,
) -> None:
    if isinstance(column, str):
        if not _is_qualified_identifier_or_quoted(column):
            raise ValueError(
                f"{context} requires qualified column names in JOIN queries"
            )
        parts = _split_qualified_identifier(column)
        if parts is not None:
            source = _parse_column_identifier(parts[0])
            name = _parse_column_identifier(parts[1])
        else:
            source, name = column.split(".", 1)
        if source not in allowed_sources or not name:
            raise ValueError(f"Invalid source reference in {context}: {source}.{name}")
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
                raise ValueError(
                    f"Invalid source reference in {context}: {source}.{name}"
                )
            return
        if column_type == "literal":
            return
        if column_type == "aggregate":
            arg = str(column.get("arg", ""))
            if arg != "*":
                if not _is_qualified_identifier_or_quoted(arg):
                    raise ValueError(
                        "Aggregate arguments in JOIN queries must be qualified column names or *"
                    )
                parts = _split_qualified_identifier(arg)
                if parts is not None:
                    source = _parse_column_identifier(parts[0])
                    name = _parse_column_identifier(parts[1])
                else:
                    source, name = arg.split(".", 1)
                if source not in allowed_sources or not name:
                    raise ValueError(
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
            _validate_join_column_reference(
                column.get("left"), allowed_sources, context
            )
            _validate_join_column_reference(
                column.get("right"), allowed_sources, context
            )
            return
        if column_type == "function":
            args = column.get("args")
            if isinstance(args, list):
                for argument in args:
                    _validate_join_column_reference(argument, allowed_sources, context)
            return
        if column_type == "cast":
            _validate_join_column_reference(
                column.get("value"), allowed_sources, context
            )
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

    raise ValueError(f"{context} requires qualified column names in JOIN queries")


def _validate_join_where_columns(where: Dict[str, Any], join_sources: set[str]) -> None:
    """Recursively validate all column references in a WHERE tree for JOIN queries."""
    for condition in where.get("conditions", []):
        _validate_join_where_node(condition, join_sources)


def _validate_join_where_node(node: Dict[str, Any], join_sources: set[str]) -> None:
    """Validate a single WHERE AST node for JOIN column references."""
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
                    raise ValueError(f"Invalid source reference in WHERE: {reference}")
        return

    # NOT node: recurse into operand
    if node.get("type") == "not":
        operand = node.get("operand")
        if isinstance(operand, dict):
            _validate_join_where_node(operand, join_sources)
        return
    # Compound or precedence-grouped node: recurse into conditions
    if "conditions" in node and node.get("type") != "not":
        for child in node["conditions"]:
            _validate_join_where_node(child, join_sources)
        return
    # Atomic condition: validate column reference
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


def _parse_order_by_item_tokens(
    tokens: list[str],
    *,
    allow_subqueries: bool = False,
    outer_sources: set[str] | None = None,
) -> dict[str, Any]:
    """Parse a single ORDER BY item like ['name', 'DESC']."""
    if not tokens:
        raise ValueError("Invalid ORDER BY clause format")

    direction = "ASC"
    expression_tokens = list(tokens)
    if expression_tokens and expression_tokens[-1].upper() in {"ASC", "DESC"}:
        direction = expression_tokens[-1].upper()
        expression_tokens = expression_tokens[:-1]

    if direction not in {"ASC", "DESC"}:
        raise ValueError("Invalid ORDER BY direction")
    if not expression_tokens:
        raise ValueError("Invalid ORDER BY clause format")
    if any(token.upper() in {"NULLS", "FIRST", "LAST"} for token in expression_tokens):
        raise ValueError(f"Unsupported SQL syntax: {' '.join(expression_tokens)}")
    if (
        len(expression_tokens) == 2
        and (
            _is_identifier_or_quoted(expression_tokens[0])
            or _is_qualified_identifier_or_quoted(expression_tokens[0])
        )
        and expression_tokens[1].isalpha()
    ):
        raise ValueError("Invalid ORDER BY direction")
    if any(token.upper() == "NULLS" for token in expression_tokens):
        nulls_index = next(
            index
            for index, token in enumerate(expression_tokens)
            if token.upper() == "NULLS"
        )
        raise ValueError(
            f"Unsupported SQL syntax: {' '.join(expression_tokens[nulls_index:])}"
        )
    if len(expression_tokens) > 1 and expression_tokens[-2].upper() in {"ASC", "DESC"}:
        raise ValueError(f"Unsupported SQL syntax: {expression_tokens[-1]}")

    if _is_case_keyword(expression_tokens[0], "CASE"):
        parsed_case, consumed = _parse_case_expression_tokens(expression_tokens, 0)
        if consumed != len(expression_tokens):
            trailing_tokens = expression_tokens[consumed:]
            if len(trailing_tokens) == 1:
                raise ValueError("Invalid ORDER BY direction")
            if trailing_tokens[0].upper() in {"ASC", "DESC"}:
                raise ValueError(
                    f"Unsupported SQL syntax: {' '.join(trailing_tokens[1:])}"
                )
            raise ValueError("Invalid ORDER BY clause format")
        result: dict[str, Any] = {
            "column": f"__expr__:{_expression_to_sql_for_order_by(parsed_case)}",
            "direction": direction,
        }
        result["__expression__"] = parsed_case
        return result

    expression_text = " ".join(expression_tokens).strip()
    parsed_aggregate = _parse_column_expression(
        expression_text,
        allow_wildcard=False,
        allow_aggregates=True,
        allow_subqueries=allow_subqueries,
        outer_sources=outer_sources,
    )
    if (
        isinstance(parsed_aggregate, dict)
        and parsed_aggregate.get("type") == "aggregate"
    ):
        return {
            "column": _aggregate_expression_to_label(parsed_aggregate),
            "direction": direction,
        }

    if _is_qualified_identifier_or_quoted(expression_text):
        parts = _split_qualified_identifier(expression_text)
        if parts is not None:
            return {
                "column": f"{_parse_column_identifier(parts[0])}.{_parse_column_identifier(parts[1])}",
                "direction": direction,
            }
        return {"column": expression_text, "direction": direction}
    if _is_identifier_or_quoted(expression_text):
        return {"column": _parse_column_identifier(expression_text), "direction": direction}

    parsed_expression = _parse_column_expression(
        expression_text,
        allow_wildcard=False,
        allow_aggregates=False,
        allow_subqueries=allow_subqueries,
        outer_sources=outer_sources,
    )
    if isinstance(parsed_expression, str):
        return {"column": parsed_expression, "direction": direction}

    result = {
        "column": f"__expr__:{_expression_to_sql_for_order_by(parsed_expression)}",
        "direction": direction,
    }
    result["__expression__"] = parsed_expression
    return result


def _parse_order_by_clause_text(
    order_part: str,
    *,
    allow_subqueries: bool = False,
    outer_sources: set[str] | None = None,
) -> list[dict[str, Any]]:
    """Parse ORDER BY clause text like 'name DESC, age ASC'."""
    order_part = _normalize_aggregate_expressions(order_part)
    tokens = _collapse_aggregate_tokens(_tokenize(order_part))
    return _parse_order_by_clause_tokens(
        tokens,
        allow_subqueries=allow_subqueries,
        outer_sources=outer_sources,
    )


def _parse_order_by_clause_tokens(
    tokens: list[str],
    *,
    allow_subqueries: bool = False,
    outer_sources: set[str] | None = None,
) -> list[dict[str, Any]]:
    """Parse ORDER BY tokens (comma-separated) into parsed items."""
    order_text = " ".join(tokens).strip()
    if not order_text:
        raise ValueError("Invalid ORDER BY clause format")

    parts = [part.strip() for part in _split_csv(order_text) if part.strip()]
    items = [
        _parse_order_by_item_tokens(
            _collapse_aggregate_tokens(_tokenize(part)),
            allow_subqueries=allow_subqueries,
            outer_sources=outer_sources,
        )
        for part in parts
    ]
    if not items:
        raise ValueError("Invalid ORDER BY clause format")
    return _OrderByClause(items)


def _parse_select(
    query: str,
    params: Optional[tuple[Any, ...]],
    *,
    _allow_subqueries: bool = True,
) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    from_index = _find_top_level_keyword_index(tokens, "FROM")
    if from_index < 0:
        raise ValueError(f"Invalid SQL query format: {query}")

    columns_token = " ".join(tokens[1:from_index]).strip()
    if not columns_token:
        raise ValueError(f"Invalid SQL query format: {query}")
    distinct = False
    if columns_token.upper().startswith("DISTINCT "):
        distinct = True
        columns_token = columns_token[len("DISTINCT ") :].strip()
        if not columns_token:
            raise ValueError("DISTINCT requires column list")

    if len(tokens) <= from_index + 1:
        raise ValueError(f"Invalid SQL query format: {query}")
    table = _parse_table_identifier(tokens[from_index + 1])
    from_entry: Dict[str, Any] = {"table": table, "alias": None, "ref": table}
    token_index = from_index + 2
    if token_index < len(tokens):
        maybe_alias = tokens[token_index]
        if maybe_alias.upper() == "AS":
            token_index += 1
            if token_index >= len(tokens):
                raise ValueError("Expected alias after AS")
            maybe_alias = tokens[token_index]
        if not _is_reserved_select_token(maybe_alias):
            from_entry["alias"] = _parse_column_identifier(maybe_alias)
            from_entry["ref"] = _parse_column_identifier(maybe_alias)
            token_index += 1

    joins: List[Dict[str, Any]] = []
    known_source_refs: set[str] = {str(from_entry["table"]), str(from_entry["ref"])}
    _known_refs: set[str] = {str(from_entry["ref"])}
    _known_tables: set[str] = {str(from_entry["table"])}
    while token_index < len(tokens):
        join_token = tokens[token_index].upper()
        join_type: Optional[str] = None
        if join_token == "JOIN":
            join_type = "INNER"
            token_index += 1
        elif join_token == "INNER":
            if (
                token_index + 1 >= len(tokens)
                or tokens[token_index + 1].upper() != "JOIN"
            ):
                raise ValueError("Unsupported SQL syntax: INNER")
            join_type = "INNER"
            token_index += 2
        elif join_token == "LEFT":
            token_index += 1
            if token_index < len(tokens) and tokens[token_index].upper() == "OUTER":
                token_index += 1
            if token_index >= len(tokens) or tokens[token_index].upper() != "JOIN":
                raise ValueError("Unsupported SQL syntax: LEFT")
            join_type = "LEFT"
            token_index += 1
        elif join_token == "RIGHT":
            token_index += 1
            if token_index < len(tokens) and tokens[token_index].upper() == "OUTER":
                token_index += 1
            if token_index >= len(tokens) or tokens[token_index].upper() != "JOIN":
                raise ValueError("Unsupported SQL syntax: RIGHT")
            join_type = "RIGHT"
            token_index += 1
        elif join_token == "FULL":
            token_index += 1
            if token_index < len(tokens) and tokens[token_index].upper() == "OUTER":
                token_index += 1
            if token_index >= len(tokens) or tokens[token_index].upper() != "JOIN":
                raise ValueError("Unsupported SQL syntax: FULL")
            join_type = "FULL"
            token_index += 1
        elif join_token == "CROSS":
            if (
                token_index + 1 >= len(tokens)
                or tokens[token_index + 1].upper() != "JOIN"
            ):
                raise ValueError("Unsupported SQL syntax: CROSS")
            join_type = "CROSS"
            token_index += 2
        else:
            break

        if token_index >= len(tokens):
            raise ValueError("Invalid JOIN clause: missing table")
        join_table = _parse_table_identifier(tokens[token_index])
        token_index += 1

        join_source: Dict[str, Any] = {
            "table": join_table,
            "alias": None,
            "ref": join_table,
        }
        if token_index < len(tokens):
            maybe_alias = tokens[token_index]
            if maybe_alias.upper() == "AS":
                token_index += 1
                if token_index >= len(tokens):
                    raise ValueError("Expected alias after AS")
                maybe_alias = tokens[token_index]
            if not _is_reserved_select_token(maybe_alias):
                join_source["alias"] = _parse_column_identifier(maybe_alias)
                join_source["ref"] = _parse_column_identifier(maybe_alias)
                token_index += 1

        right_sources = {str(join_source["table"]), str(join_source["ref"])}
        # Collision detection:
        # 1. new ref vs known refs (duplicate alias)
        # 2. new table vs known refs (table name shadows existing alias)
        # 3. new ref vs known tables (alias shadows existing table name)
        #    EXCEPT when it's the same physical table (self-join: FROM t a JOIN t b)
        new_ref = str(join_source["ref"])
        new_table = str(join_source["table"])
        collision = (_known_refs & {new_ref}) | (_known_refs & {new_table})
        # Check new ref vs known table names, but allow self-join
        if new_ref != new_table:  # has alias, so ref != table
            ref_vs_tables = _known_tables & {new_ref}
            collision |= ref_vs_tables
        else:  # no alias: ref == table
            # Only collides if table name matches existing ref (non-table)
            pass  # already covered by _known_refs check above
        if collision:
            raise ValueError(
                f"Ambiguous table reference '{collision.pop()}' in JOIN; "
                f"use distinct aliases for each table"
            )

        if join_type == "CROSS":
            if token_index < len(tokens) and tokens[token_index].upper() == "ON":
                raise ValueError("CROSS JOIN does not accept ON condition")
            joins.append(
                {
                    "type": join_type,
                    "source": join_source,
                    "on": None,
                }
            )
            known_source_refs.update(right_sources)
            _known_refs.add(new_ref)
            _known_tables.add(new_table)
            continue
        else:
            if token_index >= len(tokens) or tokens[token_index].upper() != "ON":
                raise ValueError("JOIN requires ON condition")
            token_index += 1

        on_start = token_index
        while token_index < len(tokens):
            if _is_select_clause_token(tokens[token_index]):
                break
            token_index += 1
        on_tokens = tokens[on_start:token_index]

        joins.append(
            {
                "type": join_type,
                "source": join_source,
                "on": _parse_join_on_condition(
                    on_tokens,
                    set(known_source_refs),
                    right_sources,
                ),
            }
        )
        known_source_refs.update(right_sources)
        _known_refs.add(new_ref)
        _known_tables.add(new_table)

    columns = _parse_columns(
        columns_token,
        allow_subqueries=_allow_subqueries,
        outer_sources=set(known_source_refs),
    )

    clause_tokens = tokens[token_index:]

    paren_depth = 0
    for idx, token in enumerate(clause_tokens):
        if token == "(":
            paren_depth += 1
            continue
        if token == ")":
            paren_depth -= 1
            continue
        if paren_depth > 0:
            continue
        if _is_quoted_token(token):
            continue
        token.upper()
    where = None
    group_by = None
    having = None
    order_by = None
    limit = None
    offset = None

    clause_positions = _find_clause_positions(clause_tokens)
    where_index = clause_positions.get("WHERE", -1)
    group_index = clause_positions.get("GROUP BY", -1)
    having_index = clause_positions.get("HAVING", -1)
    order_index = clause_positions.get("ORDER BY", -1)
    limit_index = clause_positions.get("LIMIT", -1)
    offset_index = clause_positions.get("OFFSET", -1)
    outer_query_sources = set(known_source_refs)

    if having_index >= 0 and group_index < 0:
        raise ValueError("HAVING requires GROUP BY")

    if where_index >= 0 and order_index >= 0 and order_index < where_index:
        raise ValueError("ORDER BY cannot appear before WHERE")
    if where_index >= 0 and limit_index >= 0 and limit_index < where_index:
        raise ValueError("LIMIT cannot appear before WHERE")
    if where_index >= 0 and offset_index >= 0 and offset_index < where_index:
        raise ValueError("OFFSET cannot appear before WHERE")
    if where_index >= 0 and group_index >= 0 and group_index < where_index:
        raise ValueError("GROUP BY cannot appear before WHERE")
    if where_index >= 0 and having_index >= 0 and having_index < where_index:
        raise ValueError("HAVING cannot appear before WHERE")
    if group_index >= 0 and having_index >= 0 and having_index < group_index:
        raise ValueError("HAVING cannot appear before GROUP BY")
    if group_index >= 0 and order_index >= 0 and order_index < group_index:
        raise ValueError("ORDER BY cannot appear before GROUP BY")
    if group_index >= 0 and limit_index >= 0 and limit_index < group_index:
        raise ValueError("LIMIT cannot appear before GROUP BY")
    if group_index >= 0 and offset_index >= 0 and offset_index < group_index:
        raise ValueError("OFFSET cannot appear before GROUP BY")
    if having_index >= 0 and order_index >= 0 and order_index < having_index:
        raise ValueError("ORDER BY cannot appear before HAVING")
    if having_index >= 0 and limit_index >= 0 and limit_index < having_index:
        raise ValueError("LIMIT cannot appear before HAVING")
    if having_index >= 0 and offset_index >= 0 and offset_index < having_index:
        raise ValueError("OFFSET cannot appear before HAVING")
    if order_index >= 0 and offset_index >= 0 and offset_index < order_index:
        raise ValueError("OFFSET cannot appear before ORDER BY")
    if limit_index >= 0 and offset_index >= 0 and offset_index < limit_index:
        raise ValueError("OFFSET cannot appear before LIMIT")

    if where_index >= 0:
        where_start = where_index + 1
        where_end_candidates = [
            idx
            for idx in [
                group_index,
                having_index,
                order_index,
                limit_index,
                offset_index,
            ]
            if idx >= 0 and idx > where_index
        ]
        where_end = (
            min(where_end_candidates) if where_end_candidates else len(clause_tokens)
        )
        where_part = " ".join(clause_tokens[where_start:where_end]).strip()
        where = _parse_where_expression(
            where_part,
            params,
            bind_params=False,
            allow_subqueries=_allow_subqueries,
            outer_sources=outer_query_sources,
        )

    if group_index >= 0:
        group_start = group_index + 2
        group_end_candidates = [
            idx
            for idx in [having_index, order_index, limit_index, offset_index]
            if idx >= 0 and idx > group_index
        ]
        group_end = (
            min(group_end_candidates) if group_end_candidates else len(clause_tokens)
        )
        group_part = " ".join(clause_tokens[group_start:group_end]).strip()
        group_columns = [col.strip() for col in _split_csv(group_part) if col.strip()]
        if not group_columns:
            raise ValueError("Invalid GROUP BY clause format")
        parsed_group_by = [
            _parse_column_expression(
                column,
                allow_wildcard=False,
                allow_aggregates=False,
            )
            for column in group_columns
        ]
        normalized_group_by: list[Any] = []
        for group_column in parsed_group_by:
            if isinstance(group_column, dict):
                source_name = group_column.get("source")
                column_name = group_column.get("name")
                if (
                    group_column.get("type") == "column"
                    and isinstance(source_name, str)
                    and isinstance(column_name, str)
                ):
                    normalized_group_by.append(f"{source_name}.{column_name}")
                    continue
                normalized_group_by.append(group_column)
                continue
            normalized_group_by.append(str(group_column))
        group_by = normalized_group_by

    if having_index >= 0:
        having_start = having_index + 1
        having_end_candidates = [
            idx
            for idx in [order_index, limit_index, offset_index]
            if idx >= 0 and idx > having_index
        ]
        having_end = (
            min(having_end_candidates) if having_end_candidates else len(clause_tokens)
        )
        having_part = " ".join(clause_tokens[having_start:having_end]).strip()
        if not having_part:
            raise ValueError("Invalid HAVING clause format")
        having_part = _normalize_aggregate_expressions(having_part)
        having = _parse_where_expression(
            having_part,
            params,
            bind_params=False,
            allow_aggregates=True,
            outer_sources=outer_query_sources,
        )

    if order_index >= 0:
        order_start = order_index + 2
        order_end_candidates = [
            idx for idx in [limit_index, offset_index] if idx >= 0 and idx > order_index
        ]
        order_end = (
            min(order_end_candidates) if order_end_candidates else len(clause_tokens)
        )
        order_part = " ".join(clause_tokens[order_start:order_end]).strip()
        order_by = _parse_order_by_clause_text(
            order_part,
            allow_subqueries=_allow_subqueries,
            outer_sources=outer_query_sources,
        )

    if limit_index >= 0:
        limit_start = limit_index + 1
        limit_end = (
            offset_index
            if offset_index >= 0 and offset_index > limit_index
            else len(clause_tokens)
        )
        limit_part = " ".join(clause_tokens[limit_start:limit_end]).strip()
        if not limit_part:
            raise ValueError("Invalid LIMIT clause format")
        limit_value = _parse_value(limit_part)
        if not isinstance(limit_value, int):
            if limit_value != "?":
                raise ValueError("LIMIT must be an integer")
        limit = limit_value

    if offset_index >= 0:
        offset_part = " ".join(clause_tokens[offset_index + 1 :]).strip()
        if not offset_part:
            raise ValueError("Invalid OFFSET clause format")
        offset_value = _parse_value(offset_part)
        if not isinstance(offset_value, int):
            if offset_value != "?":
                raise ValueError("OFFSET must be an integer")
        offset = offset_value

    consumed_indices: set[int] = set()
    for clause_name, clause_start_idx in clause_positions.items():
        if clause_name in {"GROUP BY", "ORDER BY"}:
            consumed_indices.add(clause_start_idx)
            consumed_indices.add(clause_start_idx + 1)
        else:
            consumed_indices.add(clause_start_idx)

    if where_index >= 0:
        where_start = where_index + 1
        where_end_candidates = [
            idx
            for idx in [
                group_index,
                having_index,
                order_index,
                limit_index,
                offset_index,
            ]
            if idx >= 0 and idx > where_index
        ]
        where_end = (
            min(where_end_candidates) if where_end_candidates else len(clause_tokens)
        )
        consumed_indices.update(range(where_start, where_end))

    if group_index >= 0:
        group_start = group_index + 2
        group_end_candidates = [
            idx
            for idx in [having_index, order_index, limit_index, offset_index]
            if idx >= 0 and idx > group_index
        ]
        group_end = (
            min(group_end_candidates) if group_end_candidates else len(clause_tokens)
        )
        consumed_indices.update(range(group_start, group_end))

    if having_index >= 0:
        having_start = having_index + 1
        having_end_candidates = [
            idx
            for idx in [order_index, limit_index, offset_index]
            if idx >= 0 and idx > having_index
        ]
        having_end = (
            min(having_end_candidates) if having_end_candidates else len(clause_tokens)
        )
        consumed_indices.update(range(having_start, having_end))

    if order_index >= 0:
        order_start = order_index + 2
        order_end_candidates = [
            idx for idx in [limit_index, offset_index] if idx >= 0 and idx > order_index
        ]
        order_end = (
            min(order_end_candidates) if order_end_candidates else len(clause_tokens)
        )
        consumed_indices.update(range(order_start, order_end))

    if limit_index >= 0:
        limit_start = limit_index + 1
        limit_end = (
            offset_index
            if offset_index >= 0 and offset_index > limit_index
            else len(clause_tokens)
        )
        consumed_indices.update(range(limit_start, limit_end))

    if offset_index >= 0:
        consumed_indices.update(range(offset_index + 1, len(clause_tokens)))

    unconsumed = [i for i in range(len(clause_tokens)) if i not in consumed_indices]
    if unconsumed:
        unconsumed_text = " ".join(clause_tokens[i] for i in unconsumed)
        raise ValueError(f"Unsupported SQL syntax: {unconsumed_text}")

    column_expressions: List[Any] = []
    for column in columns:
        if isinstance(column, dict) and column.get("type") == "alias":
            column_expressions.append(column.get("expression"))
        else:
            column_expressions.append(column)

    # Collect ORDER BY expression ASTs for parameter binding
    order_by_expressions: list[Any] = []
    if order_by:
        for item in order_by:
            expr = item.get("__expression__")
            if expr is not None:
                order_by_expressions.append(expr)

    has_column_placeholders = any(
        _is_placeholder(value)
        for expression in column_expressions
        for value in _expression_values_to_bind(expression)
    )
    has_order_by_placeholders = any(
        _is_placeholder(value)
        for expr in order_by_expressions
        for value in _expression_values_to_bind(expr)
    )
    if params is not None or (
        has_column_placeholders
        or (
            where
            and any(_is_placeholder(value) for value in _where_values_to_bind(where))
        )
        or (
            having
            and any(_is_placeholder(value) for value in _where_values_to_bind(having))
        )
        or has_order_by_placeholders
        or _is_placeholder(limit)
        or _is_placeholder(offset)
    ):
        values_to_bind = []
        for expression in column_expressions:
            values_to_bind.extend(_expression_values_to_bind(expression))
        if where:
            values_to_bind.extend(_where_values_to_bind(where))
        if having:
            values_to_bind.extend(_where_values_to_bind(having))
        for expr in order_by_expressions:
            values_to_bind.extend(_expression_values_to_bind(expr))
        if limit is not None:
            values_to_bind.append(limit)
        if offset is not None:
            values_to_bind.append(offset)
        bound = _bind_params(values_to_bind, params)
        consumed = 0
        for expression in column_expressions:
            consumed += _bind_expression_values(expression, bound, consumed)
        if where:
            consumed += _bind_where_conditions(where, bound, consumed)
        if having:
            consumed += _bind_where_conditions(having, bound, consumed)
        for expr in order_by_expressions:
            consumed += _bind_expression_values(expr, bound, consumed)
        if order_by:
            for item in order_by:
                expr = item.get("__expression__")
                if expr is not None:
                    item["column"] = f"__expr__:{_expression_to_sql_for_order_by(expr)}"
        if limit is not None:
            limit = bound[consumed]
            consumed += 1
        if offset is not None:
            offset = bound[consumed]
    if limit is not None and not isinstance(limit, int):
        raise ValueError("LIMIT must be an integer")
    if offset is not None and not isinstance(offset, int):
        raise ValueError("OFFSET must be an integer")
    if limit is not None:
        _validate_non_negative_pagination(limit, "LIMIT")
    if offset is not None:
        _validate_non_negative_pagination(offset, "OFFSET")

    joins_value: Optional[List[Dict[str, Any]]] = joins if joins else None
    if joins_value is not None:
        if group_by is not None:
            for group_column in group_by:
                qualified_group_column: str
                if isinstance(group_column, dict):
                    source_name = group_column.get("source")
                    column_name = group_column.get("name")
                    if (
                        group_column.get("type") == "column"
                        and isinstance(source_name, str)
                        and isinstance(column_name, str)
                    ):
                        qualified_group_column = f"{source_name}.{column_name}"
                    else:
                        raise ValueError(
                            "GROUP BY in JOIN queries requires qualified column names"
                        )
                elif isinstance(group_column, str):
                    qualified_group_column = group_column
                else:
                    raise ValueError(
                        "GROUP BY in JOIN queries requires qualified column names"
                    )

                if not _is_qualified_identifier_or_quoted(qualified_group_column):
                    raise ValueError(
                        "GROUP BY in JOIN queries requires qualified column names"
                    )

        join_sources = {
            str(from_entry["table"]),
            str(from_entry["ref"]),
        }
        for join in joins_value:
            join_sources.add(str(join["source"]["table"]))
            join_sources.add(str(join["source"]["ref"]))

        select_aliases = {
            str(column["alias"])
            for column in columns
            if isinstance(column, dict) and column.get("type") == "alias"
        }

        has_wildcard = any(isinstance(col, str) and col == "*" for col in columns)
        if has_wildcard and len(columns) > 1:
            raise ValueError(
                "SELECT * cannot be mixed with other columns in JOIN queries"
            )
        if has_wildcard:
            # Bare wildcard is valid; skip per-column validation
            pass
        else:
            for column in columns:
                expression = column
                if isinstance(column, dict) and column.get("type") == "alias":
                    expression = column.get("expression")

                if (
                    isinstance(expression, dict)
                    and expression.get("type") == "aggregate"
                ):
                    arg = str(expression.get("arg", "")).strip()
                    if arg != "*" and not _is_qualified_identifier_or_quoted(arg):
                        raise ValueError(
                            "Aggregate arguments in JOIN queries must be qualified column names or *"
                        )
                    filter_clause = expression.get("filter")
                    if isinstance(filter_clause, dict):
                        _validate_join_where_node(filter_clause, join_sources)
                    continue
                _validate_join_column_reference(expression, join_sources, "SELECT")

        if where is not None:
            _validate_join_where_columns(where, join_sources)

        if order_by is not None:
            for item in order_by:
                order_column = str(item["column"])
                if order_column in select_aliases:
                    continue
                if order_column.startswith("__expr__:"):
                    order_expression = item.get("__expression__")
                    if order_expression is not None:
                        _validate_join_column_reference(
                            order_expression, join_sources, "ORDER BY"
                        )
                    continue
                aggregate_expression = _parse_column_expression(
                    order_column,
                    allow_wildcard=False,
                    allow_aggregates=True,
                    allow_subqueries=False,
                    outer_sources=join_sources,
                )
                if (
                    isinstance(aggregate_expression, dict)
                    and aggregate_expression.get("type") == "aggregate"
                ):
                    arg = str(aggregate_expression.get("arg", "")).strip()
                    if arg != "*" and not _is_qualified_identifier_or_quoted(arg):
                        raise ValueError(
                            "Aggregate arguments in JOIN queries must be qualified column names or *"
                        )
                    filter_clause = aggregate_expression.get("filter")
                    if isinstance(filter_clause, dict):
                        _validate_join_where_node(filter_clause, join_sources)
                    continue
                parsed_order_column = _parse_qualified_column_reference(order_column)
                _validate_join_column_reference(
                    parsed_order_column, join_sources, "ORDER BY"
                )

    return {
        "action": "SELECT",
        "columns": columns,
        "table": table,
        "from": from_entry,
        "joins": joins_value,
        "where": where,
        "group_by": group_by,
        "having": having,
        "order_by": order_by,
        "limit": limit,
        "offset": offset,
        "distinct": distinct,
    }


def _bind_params(values: List[Any], params: Optional[tuple[Any, ...]]) -> List[Any]:
    if params is None:
        if any(_is_placeholder(value) for value in values):
            raise ValueError("Missing parameters for placeholders")
        return values
    bound: List[Any] = []
    param_index = 0
    for value in values:
        if _is_placeholder(value):
            if param_index >= len(params):
                raise ValueError("Not enough parameters for placeholders")
            bound.append(params[param_index])
            param_index += 1
        else:
            bound.append(value)
    if param_index < len(params):
        raise ValueError("Too many parameters for placeholders")
    return bound
