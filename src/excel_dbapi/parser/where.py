from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from ..exceptions import SqlParseError, SqlSemanticError
from ._constants import _is_placeholder
from .tokenizer import (
    _collapse_aggregate_tokens,
    _find_matching_parenthesis,
    _is_qualified_identifier_or_quoted,
    _is_quoted_token,
    _parse_column_identifier,
    _parse_value,
    _split_qualified_identifier,
    _split_csv,
    _tokenize,
)
from .expressions import (
    _aggregate_expression_to_label,
    _bind_where_conditions,
    _collect_qualified_references_from_expression,
    _expression_to_sql_for_order_by,
    _is_case_keyword,
    _normalize_aggregate_expressions,
    _parse_case_expression_tokens,
    _parse_column_expression,
    _where_values_to_bind,
)


def _parse_where_expression(
    where_part: str,
    params: Optional[tuple[Any, ...]],
    bind_params: bool = True,
    allow_aggregates: bool = False,
    allow_subqueries: bool = False,
    outer_sources: set[str] | None = None,
) -> Dict[str, Any]:
    tokens = _collapse_aggregate_tokens(_tokenize(where_part.strip()))
    if not allow_aggregates:
        paren_depth = 0
        for token in tokens:
            if token == "(":
                paren_depth += 1
                continue
            if token == ")":
                if paren_depth > 0:
                    paren_depth -= 1
                continue
            if paren_depth > 0:
                continue
            if re.fullmatch(r"(?i)(COUNT|SUM|AVG|MIN|MAX)\([^\)]+\)", token):
                raise SqlSemanticError(
                    "Aggregate functions are not allowed in WHERE clause; use HAVING instead"
                )
    if len(tokens) < 3:
        # Could be: NOT col = val (3+ tokens), (col = val) (5+ tokens)
        # Check for NOT with at least 2 following tokens
        if not (len(tokens) >= 1 and tokens[0].upper() == "NOT") and not (
            len(tokens) >= 1 and tokens[0] == "("
        ):
            raise SqlParseError("Invalid WHERE clause format")

    index = 0
    result = _parse_or_expression(
        tokens,
        index,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )
    where_expression = result[0]
    end_index = result[1]

    # Validate we consumed all tokens
    if end_index < len(tokens):
        raise SqlParseError("Invalid WHERE clause format")

    # Ensure top-level result is always in flat format for backward compatibility
    # Single atomic conditions get wrapped in {"conditions": [...], "conjunctions": []}
    if "conditions" not in where_expression and where_expression.get("type") != "not":
        where_expression = {"conditions": [where_expression], "conjunctions": []}
    elif where_expression.get("type") == "not":
        where_expression = {"conditions": [where_expression], "conjunctions": []}

    values_to_bind = _where_values_to_bind(where_expression)
    if bind_params and (
        params is not None or any(_is_placeholder(value) for value in values_to_bind)
    ):
        from .select import _bind_params
        bound = _bind_params(values_to_bind, params)
        _bind_where_conditions(where_expression, bound, 0)

    return where_expression


def _parse_or_expression(
    tokens: List[str],
    index: int,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
) -> tuple[Dict[str, Any], int]:
    """Parse an OR expression (lowest precedence)."""
    left, index = _parse_and_expression(
        tokens,
        index,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )
    conditions: List[Dict[str, Any]] = [left]
    conjunctions: List[str] = []

    while index < len(tokens) and tokens[index].upper() == "OR":
        conjunctions.append("OR")
        index += 1
        right, index = _parse_and_expression(
            tokens,
            index,
            allow_subqueries=allow_subqueries,
            allow_aggregates=allow_aggregates,
            outer_sources=outer_sources,
        )
        conditions.append(right)

    if len(conditions) == 1:
        return conditions[0], index

    return {"conditions": conditions, "conjunctions": conjunctions}, index


def _parse_and_expression(
    tokens: List[str],
    index: int,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
) -> tuple[Dict[str, Any], int]:
    """Parse an AND expression (higher precedence than OR)."""
    left, index = _parse_not_expression(
        tokens,
        index,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )
    conditions: List[Dict[str, Any]] = [left]
    conjunctions: List[str] = []

    while index < len(tokens) and tokens[index].upper() == "AND":
        # Peek: is this a BETWEEN ... AND ... ?  If so, stop.
        # BETWEEN's AND is consumed by _parse_atomic_condition already,
        # so if we're here it's a real conjunction.
        conjunctions.append("AND")
        index += 1
        right, index = _parse_not_expression(
            tokens,
            index,
            allow_subqueries=allow_subqueries,
            allow_aggregates=allow_aggregates,
            outer_sources=outer_sources,
        )
        conditions.append(right)

    if len(conditions) == 1:
        return conditions[0], index

    return {"conditions": conditions, "conjunctions": conjunctions}, index


def _parse_not_expression(
    tokens: List[str],
    index: int,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
) -> tuple[Dict[str, Any], int]:
    """Parse a NOT expression or pass through to factor."""
    if index < len(tokens) and tokens[index].upper() == "NOT":
        # Peek ahead: is this a unary NOT or part of col NOT IN/LIKE/BETWEEN?
        # Unary NOT: NOT appears at position where we expect a condition start,
        # i.e., followed by ( or a column name.  But we're called from
        # _parse_and_expression which already consumed the column,
        # so if we reach here NOT is always unary.
        index += 1
        operand, index = _parse_not_expression(
            tokens,
            index,
            allow_subqueries=allow_subqueries,
            allow_aggregates=allow_aggregates,
            outer_sources=outer_sources,
        )
        return {"type": "not", "operand": operand}, index

    return _parse_factor(
        tokens,
        index,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )


def _parse_factor(
    tokens: List[str],
    index: int,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
) -> tuple[Dict[str, Any], int]:
    """Parse a parenthesized group or an atomic condition."""
    if index >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format")

    if tokens[index].upper() == "EXISTS":
        if not allow_subqueries:
            raise SqlParseError("Subqueries are not supported in this context")
        if index + 1 >= len(tokens) or tokens[index + 1] != "(":
            raise SqlParseError("Invalid WHERE clause format: EXISTS requires '(SELECT ... )'")

        subquery_end = _find_matching_parenthesis(tokens, index + 1)
        subquery_tokens = tokens[index + 2 : subquery_end]
        if not subquery_tokens or subquery_tokens[0].upper() != "SELECT":
            raise SqlParseError("Invalid WHERE clause format: EXISTS requires SELECT subquery")

        subquery_sql = " ".join(subquery_tokens).strip()
        if "?" in _tokenize(subquery_sql):
            raise SqlParseError("Parameterized subqueries are not supported; use literal values")

        from .select import _parse_select, _detect_subquery_correlation
        parsed_subquery = _parse_select(
            subquery_sql,
            params=None,
            _allow_subqueries=True,
        )
        correlated, outer_refs = _detect_subquery_correlation(
            parsed_subquery,
            outer_sources,
        )
        return (
            {
                "type": "exists",
                "query": parsed_subquery,
                "correlated": correlated,
                "outer_refs": outer_refs,
            },
            subquery_end + 1,
        )

    # Parenthesized expression
    if tokens[index] == "(":
        index += 1  # consume '('
        expr, index = _parse_or_expression(
            tokens,
            index,
            allow_subqueries=allow_subqueries,
            allow_aggregates=allow_aggregates,
            outer_sources=outer_sources,
        )
        if index >= len(tokens) or tokens[index] != ")":
            raise SqlParseError("Invalid WHERE clause format: unmatched parenthesis")
        index += 1  # consume ')'
        # Wrap single conditions in compound so nesting is explicit
        if "conditions" not in expr and expr.get("type") != "not":
            expr = {"type": "compound", "conditions": [expr], "conjunctions": []}
        elif "conditions" in expr and "type" not in expr:
            expr = dict(expr, type="compound")
        elif expr.get("type") == "not":
            expr = {"type": "compound", "conditions": [expr], "conjunctions": []}
        return expr, index

    # Atomic condition: col OP value
    return _parse_atomic_condition(
        tokens,
        index,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )


def _is_condition_operator_start(tokens: List[str], index: int) -> bool:
    token = tokens[index]
    upper = token.upper()
    if token in {"=", "==", "!=", "<>", ">", ">=", "<", "<="}:
        return True
    if upper in {"IS", "IN", "LIKE", "ILIKE", "BETWEEN"}:
        return True
    if upper == "NOT" and index + 1 < len(tokens):
        return tokens[index + 1].upper() in {"IN", "LIKE", "ILIKE", "BETWEEN"}
    return False


def _collect_condition_expression_tokens(
    tokens: List[str],
    start_index: int,
    *,
    stop_keywords: set[str] | None = None,
    stop_at_operator: bool = False,
) -> tuple[list[str], int]:
    collected: list[str] = []
    index = start_index
    paren_depth = 0
    case_depth = 0

    while index < len(tokens):
        token = tokens[index]
        if _is_quoted_token(token):
            collected.append(token)
            index += 1
            continue

        upper = token.upper()
        if token == "(":
            paren_depth += 1
            collected.append(token)
            index += 1
            continue

        if token == ")":
            if paren_depth == 0:
                break
            paren_depth -= 1
            collected.append(token)
            index += 1
            continue

        if upper == "CASE":
            case_depth += 1
            collected.append(token)
            index += 1
            continue

        if upper == "END" and case_depth > 0:
            case_depth -= 1
            collected.append(token)
            index += 1
            continue

        if paren_depth == 0 and case_depth == 0:
            if stop_at_operator and _is_condition_operator_start(tokens, index):
                break
            if stop_keywords is not None and upper in stop_keywords:
                break

        collected.append(token)
        index += 1

    return collected, index


def _parse_condition_expression_tokens(
    expression_tokens: list[str],
    *,
    allow_aggregates: bool,
    allow_subqueries: bool,
    outer_sources: set[str] | None,
    collapse_literals: bool,
) -> Any:
    if not expression_tokens:
        raise SqlParseError("Invalid WHERE clause format")

    if _is_case_keyword(expression_tokens[0], "CASE"):
        parsed_case, consumed = _parse_case_expression_tokens(expression_tokens, 0)
        if consumed != len(expression_tokens):
            raise SqlParseError("Invalid WHERE clause format")
        return parsed_case

    if (
        expression_tokens[0] == "("
        and len(expression_tokens) > 2
        and expression_tokens[1].upper() == "SELECT"
    ):
        if not allow_subqueries:
            raise SqlParseError("Subqueries are not supported in this context")
        subquery_end = _find_matching_parenthesis(expression_tokens, 0)
        if subquery_end != len(expression_tokens) - 1:
            raise SqlParseError("Invalid WHERE clause format")
        subquery_sql = " ".join(expression_tokens[1:subquery_end]).strip()
        from .select import _parse_scalar_subquery_node
        return _parse_scalar_subquery_node(
            subquery_sql,
            outer_sources=outer_sources,
        )

    expression_text = " ".join(expression_tokens).strip()
    try:
        parsed_expression = _parse_column_expression(
            expression_text,
            allow_wildcard=False,
            allow_aggregates=allow_aggregates,
            allow_subqueries=allow_subqueries,
            outer_sources=outer_sources,
        )
    except SqlParseError as exc:
        if (
            allow_aggregates
            and "Unsupported aggregate expression" in str(exc)
            and re.fullmatch(r"(?i)(COUNT|SUM|AVG|MIN|MAX)\s*\(.+\)", expression_text)
        ):
            return _normalize_aggregate_expressions(expression_text)
        raise SqlParseError("Invalid WHERE clause format") from exc

    if (
        isinstance(parsed_expression, dict)
        and parsed_expression.get("type") == "aggregate"
    ):
        return _aggregate_expression_to_label(parsed_expression)

    if (
        collapse_literals
        and isinstance(parsed_expression, dict)
        and parsed_expression.get("type") == "literal"
    ):
        return parsed_expression.get("value")

    return parsed_expression


def _parse_condition_operand(
    tokens: List[str],
    index: int,
    *,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
) -> tuple[Any, int]:
    if index >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format")

    expression_tokens, expression_end = _collect_condition_expression_tokens(
        tokens,
        index,
        stop_at_operator=True,
    )
    if not expression_tokens:
        raise SqlParseError("Invalid WHERE clause format")

    try:
        parsed_operand = _parse_condition_expression_tokens(
            expression_tokens,
            allow_aggregates=allow_aggregates,
            allow_subqueries=allow_subqueries,
            outer_sources=outer_sources,
            collapse_literals=False,
        )
    except SqlParseError as exc:
        if "Unsupported column expression" in str(exc):
            raise SqlParseError("Invalid WHERE clause format") from exc
        raise
    return parsed_operand, expression_end


def _parse_condition_value(
    tokens: List[str],
    index: int,
    *,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
    stop_keywords: set[str] | None = None,
) -> tuple[Any, int]:
    if index >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format")

    value_stop_keywords = stop_keywords or {"AND", "OR"}
    expression_tokens, expression_end = _collect_condition_expression_tokens(
        tokens,
        index,
        stop_keywords=value_stop_keywords,
    )
    if not expression_tokens:
        raise SqlParseError("Invalid WHERE clause format")

    try:
        parsed_value = _parse_condition_expression_tokens(
            expression_tokens,
            allow_aggregates=allow_aggregates,
            allow_subqueries=allow_subqueries,
            outer_sources=outer_sources,
            collapse_literals=True,
        )
    except SqlParseError as exc:
        if "Unsupported column expression" in str(exc):
            raise SqlParseError("Invalid WHERE clause format") from exc
        raise
    return parsed_value, expression_end


def _parse_atomic_condition(
    tokens: List[str],
    index: int,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
) -> tuple[Dict[str, Any], int]:
    """Parse a single condition: col OP value."""
    if index >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format")

    column, operator_index = _parse_condition_operand(
        tokens,
        index,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )
    if operator_index >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format")

    operator_token = tokens[operator_index].upper()

    # IS NULL / IS NOT NULL
    if operator_token == "IS":
        if (
            operator_index + 1 < len(tokens)
            and tokens[operator_index + 1].upper() == "NOT"
        ):
            if (
                operator_index + 2 < len(tokens)
                and tokens[operator_index + 2].upper() == "NULL"
            ):
                return (
                    {"column": column, "operator": "IS NOT", "value": None},
                    operator_index + 3,
                )
            raise SqlParseError("Invalid WHERE clause format: expected NULL after IS NOT")
        if (
            operator_index + 1 < len(tokens)
            and tokens[operator_index + 1].upper() == "NULL"
        ):
            return (
                {"column": column, "operator": "IS", "value": None},
                operator_index + 2,
            )
        raise SqlParseError("Invalid WHERE clause format: expected NULL or NOT after IS")

    if operator_token == "NOT":
        if operator_index + 1 < len(tokens):
            next_token = tokens[operator_index + 1].upper()
            if next_token == "IN":
                return _parse_in_condition(
                    tokens,
                    column,
                    operator_index + 2,
                    allow_subqueries,
                    outer_sources=outer_sources,
                    negated=True,
                )
            if next_token in {"LIKE", "ILIKE"}:
                return _parse_like_condition(
                    tokens,
                    column,
                    operator_index + 2,
                    f"NOT {next_token}",
                    allow_subqueries=allow_subqueries,
                    allow_aggregates=allow_aggregates,
                    outer_sources=outer_sources,
                )
            if next_token == "BETWEEN":
                return _parse_between_condition(
                    tokens,
                    column,
                    operator_index + 2,
                    negated=True,
                    allow_subqueries=allow_subqueries,
                    allow_aggregates=allow_aggregates,
                    outer_sources=outer_sources,
                )
        raise SqlParseError("Invalid WHERE clause format: expected IN, LIKE, ILIKE, or BETWEEN after NOT")

    # BETWEEN
    if operator_token == "BETWEEN":
        return _parse_between_condition(
            tokens,
            column,
            operator_index + 1,
            negated=False,
            allow_subqueries=allow_subqueries,
            allow_aggregates=allow_aggregates,
            outer_sources=outer_sources,
        )

    # IN
    if operator_token == "IN":
        return _parse_in_condition(
            tokens,
            column,
            operator_index + 1,
            allow_subqueries,
            outer_sources=outer_sources,
            negated=False,
        )

    if operator_token in {"LIKE", "ILIKE"}:
        return _parse_like_condition(
            tokens,
            column,
            operator_index + 1,
            operator_token,
            allow_subqueries=allow_subqueries,
            allow_aggregates=allow_aggregates,
            outer_sources=outer_sources,
        )

    # Standard comparison: =, !=, <>, >, >=, <, <=
    if operator_index + 1 >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format")
    value, value_index = _parse_condition_value(
        tokens,
        operator_index + 1,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )
    return (
        {"column": column, "operator": tokens[operator_index], "value": value},
        value_index,
    )


def _parse_between_condition(
    tokens: List[str],
    column: Any,
    value_start: int,
    negated: bool,
    *,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
) -> tuple[Dict[str, Any], int]:
    """Parse BETWEEN low AND high."""
    low_value, and_index = _parse_condition_value(
        tokens,
        value_start,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
        stop_keywords={"AND", "OR"},
    )
    if and_index >= len(tokens) or tokens[and_index].upper() != "AND":
        raise SqlParseError("Invalid WHERE clause format: expected AND in BETWEEN clause")

    high_value, high_end = _parse_condition_value(
        tokens,
        and_index + 1,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )

    op = "NOT BETWEEN" if negated else "BETWEEN"
    return (
        {"column": column, "operator": op, "value": (low_value, high_value)},
        high_end,
    )


def _parse_like_condition(
    tokens: List[str],
    column: Any,
    value_start: int,
    operator: str,
    *,
    allow_subqueries: bool = False,
    allow_aggregates: bool = False,
    outer_sources: set[str] | None = None,
) -> tuple[Dict[str, Any], int]:
    if value_start >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format")

    value, value_index = _parse_condition_value(
        tokens,
        value_start,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
        stop_keywords={"AND", "OR", "ESCAPE"},
    )
    condition: Dict[str, Any] = {"column": column, "operator": operator, "value": value}

    if value_index < len(tokens) and tokens[value_index].upper() == "ESCAPE":
        escape_value, escape_index = _parse_condition_value(
            tokens,
            value_index + 1,
            allow_subqueries=allow_subqueries,
            allow_aggregates=allow_aggregates,
            outer_sources=outer_sources,
        )
        if not isinstance(escape_value, str) or len(escape_value) != 1:
            raise SqlParseError("Invalid WHERE clause format: ESCAPE requires a single character")
        condition["escape"] = escape_value
        value_index = escape_index

    return condition, value_index


def _parse_in_condition(
    tokens: List[str],
    column: Any,
    paren_start: int,
    allow_subqueries: bool,
    outer_sources: set[str] | None,
    negated: bool,
) -> tuple[Dict[str, Any], int]:
    """Parse IN (values...) or IN (SELECT ...)."""
    if paren_start >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format")

    in_start = paren_start
    in_end = in_start
    paren_depth = 0
    while in_end < len(tokens):
        token = tokens[in_end]
        if token == "(":
            paren_depth += 1
        elif token == ")":
            if paren_depth == 0:
                raise SqlParseError("Invalid WHERE clause format: malformed IN clause")
            paren_depth -= 1
            if paren_depth == 0:
                break
        in_end += 1
    if in_end >= len(tokens):
        raise SqlParseError("Invalid WHERE clause format: expected ')' in IN clause")

    in_values_text = " ".join(tokens[in_start : in_end + 1]).strip()
    if not in_values_text.startswith("(") or not in_values_text.endswith(")"):
        raise SqlParseError("Invalid WHERE clause format: malformed IN clause")

    raw_values = in_values_text[1:-1].strip()
    if not raw_values:
        raise SqlParseError("Invalid WHERE clause format: IN clause cannot be empty")

    op = "NOT IN" if negated else "IN"

    raw_tokens = raw_values.split()
    if raw_tokens and raw_tokens[0].upper() == "SELECT":
        if not allow_subqueries:
            raise SqlParseError("Subqueries are not supported in this context")

        if "?" in _tokenize(raw_values):
            raise SqlParseError("Parameterized subqueries are not supported; use literal values")

        # Reject JOINs inside subqueries before other checks
        subquery_tokens = _tokenize(raw_values)
        for raw_token in subquery_tokens:
            if raw_token.startswith("'") or raw_token.startswith('"'):
                continue
            if raw_token.upper() == "JOIN":
                raise SqlParseError("JOIN is not supported in subqueries")

        from .select import _parse_select, _detect_subquery_correlation
        subquery_parsed = _parse_select(
            raw_values,
            params=None,
            _allow_subqueries=False,
        )

        correlated, outer_refs = _detect_subquery_correlation(
            subquery_parsed,
            outer_sources,
        )
        if correlated:
            raise SqlParseError("Correlated subqueries are not supported")

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

        subquery_columns = subquery_parsed["columns"]
        if subquery_columns == ["*"] or len(subquery_columns) != 1:
            raise SqlParseError("Subquery in WHERE ... IN must select exactly one column")

        return (
            {
                "column": column,
                "operator": op,
                "value": {
                    "type": "subquery",
                    "query": subquery_parsed,
                    "mode": "set",
                    "correlated": False,
                    "outer_refs": [] if not outer_refs else outer_refs,
                },
            },
            in_end + 1,
        )
    else:
        parsed_values = tuple(_parse_value(token) for token in _split_csv(raw_values))
        if len(parsed_values) == 0:
            raise SqlParseError("Invalid WHERE clause format: IN clause cannot be empty")

        return (
            {
                "column": column,
                "operator": op,
                "value": parsed_values,
            },
            in_end + 1,
        )


def _is_subquery_condition(where: Dict[str, Any]) -> bool:
    """Recursively check if any condition contains a subquery value."""
    return _is_subquery_condition_node(where)


def _is_subquery_condition_node(node: Dict[str, Any]) -> bool:
    """Check a single AST node (possibly recursive) for subqueries."""
    if node.get("type") == "exists":
        return True

    # NOT node: recurse into operand
    if node.get("type") == "not":
        operand = node.get("operand")
        if isinstance(operand, dict):
            return _is_subquery_condition_node(operand)
        return False
    # Compound or precedence-grouped node: recurse into conditions
    if "conditions" in node and node.get("type") != "not":
        for child in node["conditions"]:
            if _is_subquery_condition_node(child):
                return True
        return False
    # Atomic condition: check value
    column = node.get("column")
    if isinstance(column, dict) and column.get("type") == "subquery":
        return True
    value = node.get("value")
    if isinstance(value, dict) and value.get("type") in {"subquery", "exists"}:
        return True
    return False


def _literal_to_sql_for_order_by(value: Any) -> str:
    if value is None:
        return "NULL"
    if isinstance(value, str):
        escaped = value.replace("'", "''")
        return f"'{escaped}'"
    return str(value)


def _where_operand_to_sql_for_order_by(operand: Any, *, is_column: bool) -> str:
    if isinstance(operand, dict):
        if operand.get("type") == "subquery":
            return "(SUBQUERY)"
        if operand.get("type") == "exists":
            return "EXISTS (SUBQUERY)"
        return _expression_to_sql_for_order_by(operand)
    if isinstance(operand, str):
        if is_column:
            return operand
        return _literal_to_sql_for_order_by(operand)
    return _literal_to_sql_for_order_by(operand)


def _where_to_sql_for_order_by(where: dict[str, Any]) -> str:
    node_type = where.get("type")
    if node_type == "not":
        operand = where.get("operand")
        if isinstance(operand, dict):
            return f"NOT ({_where_to_sql_for_order_by(operand)})"
        return "NOT"
    if node_type == "exists":
        return "EXISTS (SUBQUERY)"

    if "conditions" in where:
        conditions = where.get("conditions", [])
        if not conditions:
            return ""
        first = conditions[0]
        parts = [
            _where_to_sql_for_order_by(first) if isinstance(first, dict) else str(first)
        ]
        conjunctions = where.get("conjunctions", [])
        if isinstance(conjunctions, list):
            for idx, conjunction in enumerate(conjunctions):
                if idx + 1 >= len(conditions):
                    break
                parts.append(str(conjunction))
                candidate = conditions[idx + 1]
                parts.append(
                    _where_to_sql_for_order_by(candidate)
                    if isinstance(candidate, dict)
                    else str(candidate)
                )
        combined = " ".join(parts)
        if where.get("type") == "compound":
            return f"({combined})"
        return combined

    column_sql = _where_operand_to_sql_for_order_by(
        where.get("column"),
        is_column=True,
    )
    operator = str(where.get("operator", ""))
    value = where.get("value")

    if operator in {"IS", "IS NOT"}:
        return f"{column_sql} {operator} NULL"

    if operator in {"IN", "NOT IN"}:
        if isinstance(value, dict) and value.get("type") == "subquery":
            return f"{column_sql} {operator} (SUBQUERY)"
        if isinstance(value, (list, tuple)):
            values_sql = ", ".join(
                _where_operand_to_sql_for_order_by(item, is_column=False)
                for item in value
            )
        else:
            values_sql = _where_operand_to_sql_for_order_by(value, is_column=False)
        return f"{column_sql} {operator} ({values_sql})"

    if operator in {"BETWEEN", "NOT BETWEEN"} and isinstance(value, (list, tuple)):
        if len(value) != 2:
            return f"{column_sql} {operator}"
        low_sql = _where_operand_to_sql_for_order_by(value[0], is_column=False)
        high_sql = _where_operand_to_sql_for_order_by(value[1], is_column=False)
        return f"{column_sql} {operator} {low_sql} AND {high_sql}"

    value_sql = _where_operand_to_sql_for_order_by(value, is_column=False)
    return f"{column_sql} {operator} {value_sql}"


def _collect_qualified_references_from_where(where: Dict[str, Any]) -> set[str]:
    refs: set[str] = set()
    node_type = where.get("type")
    if node_type == "not":
        operand = where.get("operand")
        if isinstance(operand, dict):
            refs.update(_collect_qualified_references_from_where(operand))
        return refs

    if node_type == "exists":
        return refs

    if "conditions" in where:
        conditions = where.get("conditions")
        if isinstance(conditions, list):
            for condition in conditions:
                if isinstance(condition, dict):
                    refs.update(_collect_qualified_references_from_where(condition))
        return refs

    column = where.get("column")
    if isinstance(column, str):
        if _is_qualified_identifier_or_quoted(column):
            parts = _split_qualified_identifier(column)
            if parts is not None:
                refs.add(
                    f"{_parse_column_identifier(parts[0])}.{_parse_column_identifier(parts[1])}"
                )
            else:
                refs.add(column)
    else:
        refs.update(_collect_qualified_references_from_expression(column))

    value = where.get("value")
    if isinstance(value, dict):
        value_type = value.get("type")
        if value_type not in {"subquery", "exists"}:
            refs.update(_collect_qualified_references_from_expression(value))
    elif isinstance(value, (list, tuple)):
        for candidate in value:
            refs.update(_collect_qualified_references_from_expression(candidate))

    return refs
