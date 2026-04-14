from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from ._constants import (
    _AGGREGATE_FUNCTIONS,
    _IDENTIFIER_PATTERN,
    _SCALAR_FUNCTION_NAMES,
    _WINDOW_FUNCTIONS,
)
from .tokenizer import (
    _collapse_aggregate_tokens,
    _find_matching_parenthesis,
    _is_double_quoted_token,
    _is_identifier_or_quoted,
    _is_quoted_token,
    _is_qualified_identifier_or_quoted,
    _is_single_quoted_token,
    _parse_column_identifier,
    _parse_numeric_literal,
    _parse_value,
    _split_qualified_identifier,
    _split_csv,
    _tokenize,
    _tokenize_expression,
)


def _is_case_keyword(token: str, keyword: str) -> bool:
    return not _is_quoted_token(token) and token.upper() == keyword


def _collect_case_tokens_until(
    tokens: List[str],
    start_index: int,
    stop_keywords: set[str],
) -> tuple[List[str], int, Optional[str]]:
    collected: List[str] = []
    index = start_index
    paren_depth = 0
    case_depth = 0

    while index < len(tokens):
        token = tokens[index]
        if _is_quoted_token(token):
            collected.append(token)
            index += 1
            continue

        if token == "(":
            paren_depth += 1
            collected.append(token)
            index += 1
            continue

        if token == ")":
            if paren_depth > 0:
                paren_depth -= 1
            collected.append(token)
            index += 1
            continue

        upper = token.upper()
        if upper == "CASE":
            case_depth += 1
            collected.append(token)
            index += 1
            continue

        if upper == "END":
            if case_depth > 0:
                case_depth -= 1
                collected.append(token)
                index += 1
                continue
            if paren_depth == 0 and "END" in stop_keywords:
                return collected, index, "END"
            collected.append(token)
            index += 1
            continue

        if paren_depth == 0 and case_depth == 0 and upper in stop_keywords:
            return collected, index, upper

        collected.append(token)
        index += 1

    return collected, index, None


def _parse_case_expression_tokens(
    tokens: List[str],
    start_index: int,
) -> tuple[Dict[str, Any], int]:
    if start_index >= len(tokens) or not _is_case_keyword(tokens[start_index], "CASE"):
        raise ValueError("Invalid CASE expression")

    index = start_index + 1
    head_tokens, index, first_stop = _collect_case_tokens_until(tokens, index, {"WHEN"})
    if first_stop != "WHEN":
        raise ValueError("Invalid CASE expression: missing WHEN")

    mode = "searched"
    case_value: Any = None
    if head_tokens:
        mode = "simple"
        case_value = _parse_column_expression(
            " ".join(head_tokens),
            allow_wildcard=False,
            allow_aggregates=False,
        )

    whens: List[Dict[str, Any]] = []
    else_expression: Any = None
    closed = False

    while index < len(tokens) and _is_case_keyword(tokens[index], "WHEN"):
        index += 1
        when_tokens, index, then_stop = _collect_case_tokens_until(
            tokens, index, {"THEN"}
        )
        if then_stop != "THEN" or not when_tokens:
            raise ValueError("Invalid CASE expression: expected WHEN ... THEN ...")

        index += 1
        result_tokens, index, next_stop = _collect_case_tokens_until(
            tokens,
            index,
            {"WHEN", "ELSE", "END"},
        )
        if not result_tokens:
            raise ValueError("Invalid CASE expression: THEN requires a result")
        result_expression = _parse_column_expression(
            " ".join(result_tokens),
            allow_wildcard=False,
            allow_aggregates=False,
        )

        if mode == "searched":
            condition_text = " ".join(when_tokens).strip()
            from .where import _parse_where_expression
            condition = _parse_where_expression(
                condition_text,
                params=None,
                bind_params=False,
            )
            whens.append(
                {
                    "condition": condition,
                    "result": result_expression,
                }
            )
        else:
            match_expression = _parse_column_expression(
                " ".join(when_tokens),
                allow_wildcard=False,
                allow_aggregates=False,
            )
            whens.append(
                {
                    "match": match_expression,
                    "result": result_expression,
                }
            )

        if next_stop == "WHEN":
            continue

        if next_stop == "ELSE":
            index += 1
            else_tokens, index, end_stop = _collect_case_tokens_until(
                tokens, index, {"END"}
            )
            if end_stop != "END":
                raise ValueError("Invalid CASE expression: missing END")
            if not else_tokens:
                raise ValueError("Invalid CASE expression: ELSE requires a result")
            else_expression = _parse_column_expression(
                " ".join(else_tokens),
                allow_wildcard=False,
                allow_aggregates=False,
            )
            index += 1
            closed = True
            break

        if next_stop == "END":
            index += 1
            closed = True
            break

        if next_stop is None:
            raise ValueError("Invalid CASE expression: missing END")

        raise ValueError("Invalid CASE expression")

    if not whens:
        raise ValueError("Invalid CASE expression: missing WHEN branches")

    if not closed:
        if index < len(tokens) and _is_case_keyword(tokens[index], "END"):
            index += 1
        else:
            raise ValueError("Invalid CASE expression: missing END")

    return (
        {
            "type": "case",
            "mode": mode,
            "value": case_value,
            "whens": whens,
            "else": else_expression,
        },
        index,
    )


def _parse_case_expression(expression: str) -> Dict[str, Any]:
    tokens = _collapse_aggregate_tokens(_tokenize(expression.strip()))
    parsed_case, index = _parse_case_expression_tokens(tokens, 0)
    if index != len(tokens):
        raise ValueError("Invalid CASE expression")
    return parsed_case


def _parse_column_expression(
    expression: str,
    *,
    allow_wildcard: bool = True,
    allow_aggregates: bool = True,
    allow_subqueries: bool = False,
    outer_sources: set[str] | None = None,
) -> Any:
    expression = expression.strip()
    if not expression:
        raise ValueError("Invalid column expression")

    pretokenized = _collapse_aggregate_tokens(_tokenize(expression))
    parsed_window_or_aggregate = _parse_window_or_aggregate_expression_tokens(
        pretokenized,
        allow_aggregates=allow_aggregates,
        allow_subqueries=allow_subqueries,
        outer_sources=outer_sources,
    )
    if parsed_window_or_aggregate is not None:
        return parsed_window_or_aggregate

    if expression == "*":
        if allow_wildcard:
            return expression
        raise ValueError(
            "Unsupported column expression: wildcard is not supported here"
        )

    literal_value: Any | None = None
    parsed_numeric = _parse_numeric_literal(expression)
    if parsed_numeric is not None:
        literal_value = parsed_numeric
    elif expression.upper() == "NULL":
        literal_value = None
    elif expression == "?":
        literal_value = expression
    elif _is_single_quoted_token(expression):
        literal_value = _parse_value(expression)

    if literal_value is not None or expression.upper() == "NULL":
        return {"type": "literal", "value": literal_value}

    if _is_qualified_identifier_or_quoted(expression):
        parts = _split_qualified_identifier(expression)
        if parts is not None:
            source = _parse_column_identifier(parts[0])
            name = _parse_column_identifier(parts[1])
            return {"type": "column", "source": source, "name": name}
        source, name = expression.split(".", 1)
        return {"type": "column", "source": source, "name": name}

    if _is_identifier_or_quoted(expression):
        return _parse_column_identifier(expression)

    tokens = _tokenize_expression(expression)
    position = 0

    def _peek() -> str | None:
        if position >= len(tokens):
            return None
        return tokens[position]

    def _consume(expected: str | None = None) -> str:
        nonlocal position
        token = _peek()
        if token is None:
            raise ValueError(f"Unsupported column expression: {expression}")
        if expected is not None and token != expected:
            raise ValueError(f"Unsupported column expression: {expression}")
        position += 1
        return token

    def _parse_atom() -> Any:
        nonlocal position
        token = _peek()
        if token is None:
            raise ValueError(f"Unsupported column expression: {expression}")

        if token == "(":
            if (
                allow_subqueries
                and position + 1 < len(tokens)
                and tokens[position + 1].upper() == "SELECT"
            ):
                subquery_start = position
                subquery_end = _find_matching_parenthesis(tokens, subquery_start)
                subquery_text = " ".join(
                    tokens[subquery_start + 1 : subquery_end]
                ).strip()
                position = subquery_end + 1
                from .select import _parse_scalar_subquery_node
                return _parse_scalar_subquery_node(
                    subquery_text,
                    outer_sources=outer_sources,
                )

            _consume("(")
            parsed = _parse_expression_internal()
            if _peek() != ")":
                raise ValueError(f"Unsupported column expression: {expression}")
            _consume(")")
            return parsed

        if _is_case_keyword(token, "CASE"):
            remaining_tokens = tokens[position:]
            parsed_case, consumed = _parse_case_expression_tokens(remaining_tokens, 0)
            position += consumed
            return parsed_case

        if (
            re.fullmatch(_IDENTIFIER_PATTERN, token)
            and position + 1 < len(tokens)
            and tokens[position + 1] == "("
        ):
            function_name = token.upper()
            _consume()
            _consume("(")

            if function_name == "CAST":
                cast_value = _parse_expression_internal()
                as_token = _peek()
                if as_token is None or as_token.upper() != "AS":
                    raise ValueError(f"Unsupported column expression: {expression}")
                _consume()
                target_type = _consume()
                if not re.fullmatch(_IDENTIFIER_PATTERN, target_type):
                    raise ValueError(f"Unsupported column expression: {expression}")
                if _peek() != ")":
                    raise ValueError(f"Unsupported column expression: {expression}")
                _consume(")")
                return {
                    "type": "cast",
                    "value": cast_value,
                    "target_type": target_type.upper(),
                }

            if function_name not in _SCALAR_FUNCTION_NAMES:
                if function_name in _AGGREGATE_FUNCTIONS:
                    raise ValueError(
                        f"Unsupported function: {token}. "
                        f"Unsupported aggregate expression: {expression}. "
                        "Only bare column names and * are supported"
                    )
                raise ValueError(f"Unsupported function: {token}")

            arguments: list[Any] = []
            if _peek() == ")":
                _consume(")")
            else:
                while True:
                    arguments.append(_parse_expression_internal())
                    if _peek() != ",":
                        break
                    _consume(",")
                if _peek() != ")":
                    raise ValueError(f"Unsupported column expression: {expression}")
                _consume(")")

            return {"type": "function", "name": function_name, "args": arguments}

        _consume()

        numeric = _parse_numeric_literal(token)
        if numeric is not None:
            return {"type": "literal", "value": numeric}

        if token.upper() == "NULL" or token == "?" or _is_single_quoted_token(token):
            return {"type": "literal", "value": _parse_value(token)}

        if _is_double_quoted_token(token):
            return _parse_column_identifier(token)

        if _is_qualified_identifier_or_quoted(token):
            parts = _split_qualified_identifier(token)
            if parts is not None:
                source = _parse_column_identifier(parts[0])
                name = _parse_column_identifier(parts[1])
                return {"type": "column", "source": source, "name": name}
            source, name = token.split(".", 1)
            return {"type": "column", "source": source, "name": name}

        if _is_identifier_or_quoted(token):
            return _parse_column_identifier(token)

        if re.fullmatch(
            r"(?i)(COUNT|SUM|AVG|MIN|MAX)\s*\(\s*(DISTINCT\s+)?([^\)]+?)\s*\)",
            token,
        ):
            raise ValueError(
                "Unsupported column expression: aggregate functions cannot be used inside arithmetic expressions"
            )

        raise ValueError(
            f"Unsupported column expression: {expression}. "
            "Only bare column names, qualified column names, numeric literals, "
            "quoted literals, CASE expressions, CAST, scalar functions, "
            "arithmetic operators (+, -, *, /, ||), and aggregate functions are supported"
        )

    def _parse_factor() -> Any:
        token = _peek()
        if token == "-":
            _consume("-")
            return {"type": "unary_op", "op": "-", "operand": _parse_factor()}
        return _parse_atom()

    def _parse_term() -> Any:
        left = _parse_factor()
        while True:
            op = _peek()
            if op not in {"*", "/"}:
                break
            _consume(op)
            right = _parse_factor()
            left = {"type": "binary_op", "op": op, "left": left, "right": right}
        return left

    def _parse_additive() -> Any:
        left = _parse_term()
        while True:
            op = _peek()
            if op not in {"+", "-"}:
                break
            _consume(op)
            right = _parse_term()
            left = {"type": "binary_op", "op": op, "left": left, "right": right}
        return left

    def _parse_expression_internal() -> Any:
        left = _parse_additive()
        while True:
            op = _peek()
            if op != "||":
                break
            _consume(op)
            right = _parse_additive()
            left = {"type": "binary_op", "op": op, "left": left, "right": right}
        return left

    parsed_expression = _parse_expression_internal()
    if _peek() is not None:
        raise ValueError(f"Unsupported column expression: {expression}")
    return parsed_expression


def _aggregate_expression_to_label(aggregate: dict[str, Any]) -> str:
    func = str(aggregate.get("func", "")).upper()
    arg = str(aggregate.get("arg", "")).strip()
    aggregate_sql = (
        f"{func}(DISTINCT {arg})" if aggregate.get("distinct") else f"{func}({arg})"
    )
    filter_clause = aggregate.get("filter")
    if isinstance(filter_clause, dict):
        from .where import _where_to_sql_for_order_by
        filter_sql = _where_to_sql_for_order_by(filter_clause)
        aggregate_sql = f"{aggregate_sql} FILTER (WHERE {filter_sql})"
    return aggregate_sql


def _parse_aggregate_token(token: str) -> dict[str, Any] | None:
    match = re.fullmatch(
        r"(?i)(COUNT|SUM|AVG|MIN|MAX)\s*\(\s*(DISTINCT\s+)?([^\)]+?)\s*\)",
        token,
    )
    if not match:
        return None

    func = match.group(1).upper()
    distinct_modifier = match.group(2)
    arg = match.group(3).strip()
    if not arg:
        raise ValueError("Invalid aggregate expression")
    if distinct_modifier:
        if func != "COUNT":
            raise ValueError("DISTINCT is only supported with COUNT")
        if arg == "*":
            raise ValueError("Invalid aggregate expression")
        if not (_is_identifier_or_quoted(arg) or _is_qualified_identifier_or_quoted(arg)):
            raise ValueError(
                f"Unsupported aggregate expression: COUNT(DISTINCT {arg}). "
                "Only bare and qualified column names are supported with DISTINCT"
            )

    if arg == "*" and func != "COUNT":
        raise ValueError(f"{func} does not support *")
    if arg != "*" and not (
        _is_identifier_or_quoted(arg) or _is_qualified_identifier_or_quoted(arg)
    ):
        raise ValueError(
            f"Unsupported aggregate expression: {func}({arg}). "
            "Only bare column names and * are supported"
        )

    aggregate: dict[str, Any] = {"type": "aggregate", "func": func, "arg": arg}
    if distinct_modifier:
        aggregate["distinct"] = True
    return aggregate


def _parse_filter_clause_tokens(
    tokens: list[str],
    start_index: int,
    *,
    allow_subqueries: bool,
    outer_sources: set[str] | None,
) -> tuple[dict[str, Any] | None, int]:
    if start_index >= len(tokens) or tokens[start_index].upper() != "FILTER":
        return None, start_index

    if start_index + 2 >= len(tokens) or tokens[start_index + 1] != "(":
        raise ValueError("Invalid FILTER clause")
    if tokens[start_index + 2].upper() != "WHERE":
        raise ValueError("Invalid FILTER clause: expected WHERE")

    filter_end = _find_matching_parenthesis(tokens, start_index + 1)
    filter_tokens = tokens[start_index + 3 : filter_end]
    if not filter_tokens:
        raise ValueError("Invalid FILTER clause")

    filter_text = " ".join(filter_tokens).strip()
    from .where import _parse_where_expression
    filter_clause = _parse_where_expression(
        filter_text,
        params=None,
        bind_params=False,
        allow_aggregates=False,
        allow_subqueries=allow_subqueries,
        outer_sources=outer_sources,
    )
    return filter_clause, filter_end + 1


def _find_top_level_window_clause_index(
    tokens: list[str],
    start_index: int,
) -> int:
    depth = 0
    index = start_index
    while index < len(tokens):
        token = tokens[index]
        upper = token.upper()
        if token == "(":
            depth += 1
            index += 1
            continue
        if token == ")":
            if depth > 0:
                depth -= 1
            index += 1
            continue
        if depth == 0 and (
            (
                upper == "ORDER"
                and index + 1 < len(tokens)
                and tokens[index + 1].upper() == "BY"
            )
            or upper == "ROWS"
        ):
            return index
        index += 1
    return len(tokens)


def _parse_window_spec_tokens(
    tokens: list[str],
    start_index: int,
    *,
    outer_sources: set[str] | None,
) -> tuple[list[Any], list[dict[str, Any]], int]:
    if start_index >= len(tokens) or tokens[start_index].upper() != "OVER":
        raise ValueError("Invalid window function: missing OVER")
    if start_index + 1 >= len(tokens) or tokens[start_index + 1] != "(":
        raise ValueError("Invalid window specification")

    spec_end = _find_matching_parenthesis(tokens, start_index + 1)
    spec_tokens = tokens[start_index + 2 : spec_end]

    partition_by: list[Any] = []
    order_by: list[dict[str, Any]] = []
    index = 0

    if index < len(spec_tokens) and spec_tokens[index].upper() == "PARTITION":
        if index + 1 >= len(spec_tokens) or spec_tokens[index + 1].upper() != "BY":
            raise ValueError(
                "Invalid window specification: expected BY after PARTITION"
            )
        partition_start = index + 2
        partition_end = _find_top_level_window_clause_index(
            spec_tokens, partition_start
        )
        partition_text = " ".join(spec_tokens[partition_start:partition_end]).strip()
        if not partition_text:
            raise ValueError(
                "Invalid window specification: PARTITION BY requires expression"
            )
        partition_by = [
            _parse_column_expression(
                part,
                allow_wildcard=False,
                allow_aggregates=False,
                allow_subqueries=False,
                outer_sources=outer_sources,
            )
            for part in _split_csv(partition_text)
            if part.strip()
        ]
        if not partition_by:
            raise ValueError(
                "Invalid window specification: PARTITION BY requires expression"
            )
        index = partition_end

    if index < len(spec_tokens) and spec_tokens[index].upper() == "ORDER":
        if index + 1 >= len(spec_tokens) or spec_tokens[index + 1].upper() != "BY":
            raise ValueError("Invalid window specification: expected BY after ORDER")
        order_start = index + 2
        order_end = _find_top_level_window_clause_index(spec_tokens, order_start)
        order_text = " ".join(spec_tokens[order_start:order_end]).strip()
        if not order_text:
            raise ValueError(
                "Invalid window specification: ORDER BY requires expression"
            )
        from .select import _parse_order_by_clause_text
        order_by = list(
            _parse_order_by_clause_text(
                order_text,
                allow_subqueries=False,
                outer_sources=outer_sources,
            )
        )
        index = order_end

    if index < len(spec_tokens):
        frame_tokens = [token.upper() for token in spec_tokens[index:]]
        if frame_tokens != [
            "ROWS",
            "BETWEEN",
            "UNBOUNDED",
            "PRECEDING",
            "AND",
            "CURRENT",
            "ROW",
        ]:
            raise ValueError("Unsupported window frame specification")

    return partition_by, order_by, spec_end + 1


def _parse_window_or_aggregate_expression_tokens(
    tokens: list[str],
    *,
    allow_aggregates: bool,
    allow_subqueries: bool,
    outer_sources: set[str] | None,
) -> Any | None:
    if not tokens:
        return None

    aggregate_expression = _parse_aggregate_token(tokens[0])
    if aggregate_expression is not None:
        index = 1
        filter_clause, index = _parse_filter_clause_tokens(
            tokens,
            index,
            allow_subqueries=allow_subqueries,
            outer_sources=outer_sources,
        )
        if filter_clause is not None:
            aggregate_expression["filter"] = filter_clause

        if index < len(tokens) and tokens[index].upper() == "OVER":
            partition_by, order_by, index = _parse_window_spec_tokens(
                tokens,
                index,
                outer_sources=outer_sources,
            )
            if index != len(tokens):
                raise ValueError("Unsupported column expression")
            window_function: dict[str, Any] = {
                "type": "window_function",
                "func": str(aggregate_expression["func"]),
                "args": [str(aggregate_expression["arg"])],
                "partition_by": partition_by,
                "order_by": order_by,
            }
            if aggregate_expression.get("distinct"):
                window_function["distinct"] = True
            if filter_clause is not None:
                window_function["filter"] = filter_clause
            return window_function

        if index == len(tokens):
            if not allow_aggregates:
                raise ValueError(
                    "Unsupported column expression: aggregate functions are not supported here"
                )
            return aggregate_expression

    if (
        len(tokens) >= 4
        and re.fullmatch(_IDENTIFIER_PATTERN, tokens[0])
        and tokens[1] == "("
    ):
        function_name = tokens[0].upper()
        if function_name in _WINDOW_FUNCTIONS:
            function_end = _find_matching_parenthesis(tokens, 1)
            if function_end != 2:
                raise ValueError(f"{function_name} does not accept arguments")
            if (
                function_end + 1 >= len(tokens)
                or tokens[function_end + 1].upper() != "OVER"
            ):
                return None

            partition_by, order_by, next_index = _parse_window_spec_tokens(
                tokens,
                function_end + 1,
                outer_sources=outer_sources,
            )
            if next_index != len(tokens):
                raise ValueError("Unsupported column expression")
            return {
                "type": "window_function",
                "func": function_name,
                "args": [],
                "partition_by": partition_by,
                "order_by": order_by,
            }

    return None




def _normalize_aggregate_expressions(text: str) -> str:
    return " ".join(_collapse_aggregate_tokens(_tokenize(text)))


def _expression_values_to_bind(expression: Any) -> List[Any]:
    if isinstance(expression, dict):
        expression_type = expression.get("type")
        if expression_type == "alias":
            return _expression_values_to_bind(expression.get("expression"))
        if expression_type == "literal":
            return [expression.get("value")]
        if expression_type == "unary_op":
            return _expression_values_to_bind(expression.get("operand"))
        if expression_type == "binary_op":
            values = _expression_values_to_bind(expression.get("left"))
            values.extend(_expression_values_to_bind(expression.get("right")))
            return values
        if expression_type == "function":
            function_values: List[Any] = []
            args = expression.get("args")
            if isinstance(args, list):
                for argument in args:
                    function_values.extend(_expression_values_to_bind(argument))
            return function_values
        if expression_type == "cast":
            return _expression_values_to_bind(expression.get("value"))
        if expression_type == "aggregate":
            aggregate_values: List[Any] = []
            filter_clause = expression.get("filter")
            if isinstance(filter_clause, dict):
                aggregate_values.extend(_where_values_to_bind(filter_clause))
            return aggregate_values
        if expression_type == "window_function":
            window_values: List[Any] = []
            args = expression.get("args")
            if isinstance(args, list):
                for argument in args:
                    window_values.extend(_expression_values_to_bind(argument))

            partition_by = expression.get("partition_by")
            if isinstance(partition_by, list):
                for partition_expression in partition_by:
                    window_values.extend(
                        _expression_values_to_bind(partition_expression)
                    )

            order_by = expression.get("order_by")
            if isinstance(order_by, list):
                for item in order_by:
                    if not isinstance(item, dict):
                        continue
                    order_expression = item.get("__expression__")
                    if order_expression is not None:
                        window_values.extend(
                            _expression_values_to_bind(order_expression)
                        )

            filter_clause = expression.get("filter")
            if isinstance(filter_clause, dict):
                window_values.extend(_where_values_to_bind(filter_clause))
            return window_values
        if expression_type == "case":
            case_values: List[Any] = []
            mode = str(expression.get("mode", ""))
            if mode == "simple":
                case_values.extend(_expression_values_to_bind(expression.get("value")))
            whens = expression.get("whens")
            if isinstance(whens, list):
                for when in whens:
                    if not isinstance(when, dict):
                        continue
                    if mode == "searched":
                        condition = when.get("condition")
                        if isinstance(condition, dict):
                            case_values.extend(_where_values_to_bind(condition))
                    else:
                        case_values.extend(
                            _expression_values_to_bind(when.get("match"))
                        )
                    case_values.extend(_expression_values_to_bind(when.get("result")))
            case_values.extend(_expression_values_to_bind(expression.get("else")))
            return case_values
    return []


def _bind_expression_values(
    expression: Any,
    bound_values: List[Any],
    offset: int,
) -> int:
    if not isinstance(expression, dict):
        return 0

    expression_type = expression.get("type")
    if expression_type == "alias":
        return _bind_expression_values(
            expression.get("expression"), bound_values, offset
        )

    if expression_type == "literal":
        expression["value"] = bound_values[offset]
        return 1

    if expression_type == "unary_op":
        return _bind_expression_values(expression.get("operand"), bound_values, offset)

    if expression_type == "binary_op":
        consumed = _bind_expression_values(expression.get("left"), bound_values, offset)
        consumed += _bind_expression_values(
            expression.get("right"),
            bound_values,
            offset + consumed,
        )
        return consumed

    if expression_type == "function":
        consumed = 0
        args = expression.get("args")
        if isinstance(args, list):
            for argument in args:
                consumed += _bind_expression_values(
                    argument,
                    bound_values,
                    offset + consumed,
                )
        return consumed

    if expression_type == "cast":
        return _bind_expression_values(expression.get("value"), bound_values, offset)

    if expression_type == "aggregate":
        filter_clause = expression.get("filter")
        if isinstance(filter_clause, dict):
            return _bind_where_conditions(filter_clause, bound_values, offset)
        return 0

    if expression_type == "window_function":
        consumed = 0

        args = expression.get("args")
        if isinstance(args, list):
            for argument in args:
                consumed += _bind_expression_values(
                    argument,
                    bound_values,
                    offset + consumed,
                )

        partition_by = expression.get("partition_by")
        if isinstance(partition_by, list):
            for partition_expression in partition_by:
                consumed += _bind_expression_values(
                    partition_expression,
                    bound_values,
                    offset + consumed,
                )

        order_by = expression.get("order_by")
        if isinstance(order_by, list):
            for item in order_by:
                if not isinstance(item, dict):
                    continue
                order_expression = item.get("__expression__")
                if order_expression is None:
                    continue
                consumed += _bind_expression_values(
                    order_expression,
                    bound_values,
                    offset + consumed,
                )

        filter_clause = expression.get("filter")
        if isinstance(filter_clause, dict):
            consumed += _bind_where_conditions(
                filter_clause,
                bound_values,
                offset + consumed,
            )
        return consumed

    if expression_type == "case":
        consumed = 0
        mode = str(expression.get("mode", ""))
        if mode == "simple":
            consumed += _bind_expression_values(
                expression.get("value"), bound_values, offset + consumed
            )

        whens = expression.get("whens")
        if isinstance(whens, list):
            for when in whens:
                if not isinstance(when, dict):
                    continue
                if mode == "searched":
                    condition = when.get("condition")
                    if isinstance(condition, dict):
                        consumed += _bind_where_conditions(
                            condition,
                            bound_values,
                            offset + consumed,
                        )
                else:
                    consumed += _bind_expression_values(
                        when.get("match"),
                        bound_values,
                        offset + consumed,
                    )
                consumed += _bind_expression_values(
                    when.get("result"),
                    bound_values,
                    offset + consumed,
                )

        consumed += _bind_expression_values(
            expression.get("else"),
            bound_values,
            offset + consumed,
        )
        return consumed

    return 0


def _values_to_bind_from_condition(condition: Dict[str, Any]) -> List[Any]:
    # Handle compound/not nodes recursively
    node_type = condition.get("type")
    if node_type == "not":
        return _where_values_to_bind(condition["operand"])
    if node_type == "compound" or "conditions" in condition:
        return _where_values_to_bind(condition)
    if node_type == "exists":
        return []

    values: List[Any] = []
    column_operand = condition.get("column")
    if isinstance(column_operand, dict):
        values.extend(_expression_values_to_bind(column_operand))

    operator = str(condition["operator"]).upper()
    if operator in {"IS", "IS NOT"}:
        return values

    value = condition["value"]
    if operator in {"IN", "NOT IN"}:
        if isinstance(value, dict) and value.get("type") == "subquery":
            return values
        for item in value:
            if isinstance(item, dict):
                values.extend(_expression_values_to_bind(item))
            else:
                values.append(item)
        return values

    if operator in {"BETWEEN", "NOT BETWEEN"}:
        low, high = value
        if isinstance(low, dict):
            values.extend(_expression_values_to_bind(low))
        else:
            values.append(low)
        if isinstance(high, dict):
            values.extend(_expression_values_to_bind(high))
        else:
            values.append(high)
        return values

    if isinstance(value, dict):
        values.extend(_expression_values_to_bind(value))
        return values

    values.append(value)
    return values


def _apply_bound_values_to_condition(
    condition: Dict[str, Any], bound_values: List[Any], offset: int
) -> int:
    # Handle compound/not nodes recursively
    node_type = condition.get("type")
    if node_type == "not":
        return _bind_where_conditions(condition["operand"], bound_values, offset)
    if node_type == "compound" or "conditions" in condition:
        return _bind_where_conditions(condition, bound_values, offset)
    if node_type == "exists":
        return 0

    consumed = 0
    column_operand = condition.get("column")
    if isinstance(column_operand, dict):
        consumed += _bind_expression_values(
            column_operand, bound_values, offset + consumed
        )

    operator = str(condition["operator"]).upper()
    if operator in {"IS", "IS NOT"}:
        return consumed

    value = condition["value"]
    if operator in {"IN", "NOT IN"}:
        if isinstance(value, dict) and value.get("type") == "subquery":
            return consumed
        updated_values: List[Any] = []
        for item in value:
            if isinstance(item, dict):
                used = _bind_expression_values(item, bound_values, offset + consumed)
                consumed += used
                updated_values.append(item)
            else:
                updated_values.append(bound_values[offset + consumed])
                consumed += 1
        condition["value"] = tuple(updated_values)
        return consumed

    if operator in {"BETWEEN", "NOT BETWEEN"}:
        low, high = value
        if isinstance(low, dict):
            consumed += _bind_expression_values(low, bound_values, offset + consumed)
        else:
            low = bound_values[offset + consumed]
            consumed += 1
        if isinstance(high, dict):
            consumed += _bind_expression_values(high, bound_values, offset + consumed)
        else:
            high = bound_values[offset + consumed]
            consumed += 1
        condition["value"] = (low, high)
        return consumed

    if isinstance(value, dict):
        consumed += _bind_expression_values(value, bound_values, offset + consumed)
        return consumed

    condition["value"] = bound_values[offset + consumed]
    consumed += 1
    return consumed


def _where_values_to_bind(where: Dict[str, Any]) -> List[Any]:
    node_type = where.get("type")
    if node_type == "not":
        return _where_values_to_bind(where["operand"])
    if "conditions" not in where:
        # Single atomic condition
        return _values_to_bind_from_condition(where)
    values: List[Any] = []
    for condition in where["conditions"]:
        values.extend(_values_to_bind_from_condition(condition))
    return values


def _bind_where_conditions(
    where: Dict[str, Any], bound_values: List[Any], offset: int
) -> int:
    node_type = where.get("type")
    if node_type == "not":
        return _bind_where_conditions(where["operand"], bound_values, offset)
    if "conditions" not in where:
        # Single atomic condition
        return _apply_bound_values_to_condition(where, bound_values, offset)
    consumed = 0
    for condition in where["conditions"]:
        used = _apply_bound_values_to_condition(
            condition, bound_values, offset + consumed
        )
        consumed += used
    return consumed


def _collect_qualified_references_from_expression(expression: Any) -> set[str]:
    refs: set[str] = set()
    if isinstance(expression, str):
        if _is_qualified_identifier_or_quoted(expression):
            parts = _split_qualified_identifier(expression)
            if parts is not None:
                refs.add(
                    f"{_parse_column_identifier(parts[0])}.{_parse_column_identifier(parts[1])}"
                )
            else:
                refs.add(expression)
        return refs

    if not isinstance(expression, dict):
        return refs

    expression_type = expression.get("type")
    if expression_type == "alias":
        refs.update(
            _collect_qualified_references_from_expression(expression.get("expression"))
        )
        return refs

    if expression_type == "column":
        source = expression.get("source", expression.get("table"))
        name = expression.get("name")
        if isinstance(source, str) and isinstance(name, str):
            refs.add(f"{source}.{name}")
        return refs

    if expression_type == "aggregate":
        arg = expression.get("arg")
        if isinstance(arg, str) and _is_qualified_identifier_or_quoted(arg):
            parts = _split_qualified_identifier(arg)
            if parts is not None:
                refs.add(
                    f"{_parse_column_identifier(parts[0])}.{_parse_column_identifier(parts[1])}"
                )
            else:
                refs.add(arg)
        filter_clause = expression.get("filter")
        if isinstance(filter_clause, dict):
            from .where import _collect_qualified_references_from_where
            refs.update(_collect_qualified_references_from_where(filter_clause))
        return refs

    if expression_type == "window_function":
        args = expression.get("args")
        if isinstance(args, list):
            for argument in args:
                refs.update(_collect_qualified_references_from_expression(argument))

        partition_by = expression.get("partition_by")
        if isinstance(partition_by, list):
            for partition_expression in partition_by:
                refs.update(
                    _collect_qualified_references_from_expression(partition_expression)
                )

        order_by = expression.get("order_by")
        if isinstance(order_by, list):
            for item in order_by:
                if not isinstance(item, dict):
                    continue
                order_expression = item.get("__expression__")
                if order_expression is not None:
                    refs.update(
                        _collect_qualified_references_from_expression(order_expression)
                    )
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

        filter_clause = expression.get("filter")
        if isinstance(filter_clause, dict):
            from .where import _collect_qualified_references_from_where
            refs.update(_collect_qualified_references_from_where(filter_clause))
        return refs

    if expression_type == "literal":
        return refs

    if expression_type == "unary_op":
        refs.update(
            _collect_qualified_references_from_expression(expression.get("operand"))
        )
        return refs

    if expression_type == "binary_op":
        refs.update(
            _collect_qualified_references_from_expression(expression.get("left"))
        )
        refs.update(
            _collect_qualified_references_from_expression(expression.get("right"))
        )
        return refs

    if expression_type == "function":
        args = expression.get("args")
        if isinstance(args, list):
            for argument in args:
                refs.update(_collect_qualified_references_from_expression(argument))
        return refs

    if expression_type == "cast":
        refs.update(
            _collect_qualified_references_from_expression(expression.get("value"))
        )
        return refs

    if expression_type == "case":
        refs.update(
            _collect_qualified_references_from_expression(expression.get("value"))
        )
        whens = expression.get("whens")
        mode = str(expression.get("mode", ""))
        if isinstance(whens, list):
            for when_branch in whens:
                if not isinstance(when_branch, dict):
                    continue
                if mode == "searched":
                    condition = when_branch.get("condition")
                    if isinstance(condition, dict):
                        from .where import _collect_qualified_references_from_where
                        refs.update(_collect_qualified_references_from_where(condition))
                else:
                    refs.update(
                        _collect_qualified_references_from_expression(
                            when_branch.get("match")
                        )
                    )
                refs.update(
                    _collect_qualified_references_from_expression(
                        when_branch.get("result")
                    )
                )
        refs.update(
            _collect_qualified_references_from_expression(expression.get("else"))
        )
        return refs

    return refs


def _expression_to_sql_for_order_by(expression: Any) -> str:
    from .where import _literal_to_sql_for_order_by, _where_to_sql_for_order_by

    if isinstance(expression, dict):
        expression_type = expression.get("type")
        if expression_type == "alias":
            return _expression_to_sql_for_order_by(expression.get("expression"))
        if expression_type == "column":
            return f"{expression['source']}.{expression['name']}"
        if expression_type == "aggregate":
            func = str(expression.get("func", "")).upper()
            arg = str(expression.get("arg", "")).strip()
            aggregate_sql = (
                f"{func}(DISTINCT {arg})"
                if expression.get("distinct")
                else f"{func}({arg})"
            )
            filter_clause = expression.get("filter")
            if isinstance(filter_clause, dict):
                filter_sql = _where_to_sql_for_order_by(filter_clause)
                aggregate_sql = f"{aggregate_sql} FILTER (WHERE {filter_sql})"
            return aggregate_sql
        if expression_type == "window_function":
            func = str(expression.get("func", "")).upper()
            args = expression.get("args")
            args_list = args if isinstance(args, list) else []
            args_sql_parts = [
                _expression_to_sql_for_order_by(argument)
                if isinstance(argument, dict)
                else str(argument)
                for argument in args_list
            ]
            args_sql = ", ".join(args_sql_parts)
            if expression.get("distinct") and args_sql:
                args_sql = f"DISTINCT {args_sql}"

            function_sql = f"{func}({args_sql})"
            filter_clause = expression.get("filter")
            if isinstance(filter_clause, dict):
                filter_sql = _where_to_sql_for_order_by(filter_clause)
                function_sql = f"{function_sql} FILTER (WHERE {filter_sql})"

            spec_parts: list[str] = []
            partition_by = expression.get("partition_by")
            if isinstance(partition_by, list) and partition_by:
                partition_sql = ", ".join(
                    _expression_to_sql_for_order_by(partition_expression)
                    for partition_expression in partition_by
                )
                spec_parts.append(f"PARTITION BY {partition_sql}")

            order_by = expression.get("order_by")
            if isinstance(order_by, list) and order_by:
                order_parts: list[str] = []
                for order_item in order_by:
                    if not isinstance(order_item, dict):
                        continue
                    order_expression = order_item.get("__expression__")
                    if order_expression is not None:
                        order_column_sql = _expression_to_sql_for_order_by(
                            order_expression
                        )
                    else:
                        order_column_sql = str(order_item.get("column", ""))
                        if order_column_sql.startswith("__expr__:"):
                            order_column_sql = order_column_sql[len("__expr__:") :]
                    direction = str(order_item.get("direction", "ASC")).upper()
                    order_parts.append(f"{order_column_sql} {direction}")
                if order_parts:
                    spec_parts.append("ORDER BY " + ", ".join(order_parts))

            if spec_parts:
                return f"{function_sql} OVER ({' '.join(spec_parts)})"
            return f"{function_sql} OVER ()"
        if expression_type == "literal":
            return _literal_to_sql_for_order_by(expression.get("value"))
        if expression_type == "unary_op":
            operand_sql = _expression_to_sql_for_order_by(expression.get("operand"))
            return f"-{operand_sql}"
        if expression_type == "binary_op":
            left_sql = _expression_to_sql_for_order_by(expression.get("left"))
            right_sql = _expression_to_sql_for_order_by(expression.get("right"))
            return f"({left_sql} {expression['op']} {right_sql})"
        if expression_type == "function":
            args = expression.get("args", [])
            args_sql = ", ".join(_expression_to_sql_for_order_by(arg) for arg in args)
            return f"{expression['name']}({args_sql})"
        if expression_type == "cast":
            value_sql = _expression_to_sql_for_order_by(expression.get("value"))
            target_type = str(expression.get("target_type", ""))
            return f"CAST({value_sql} AS {target_type})"
        if expression_type == "subquery":
            return "(SUBQUERY)"
        if expression_type == "case":
            parts: list[str] = ["CASE"]
            mode = expression.get("mode", "searched")
            if mode == "simple" and expression.get("value") is not None:
                parts.append(_expression_to_sql_for_order_by(expression["value"]))
            whens = expression.get("whens", [])
            if isinstance(whens, list):
                for when_branch in whens:
                    if not isinstance(when_branch, dict):
                        continue
                    if mode == "searched":
                        condition = when_branch.get("condition")
                        condition_sql = ""
                        if isinstance(condition, dict):
                            condition_sql = _where_to_sql_for_order_by(condition)
                        parts.append(f"WHEN {condition_sql} THEN")
                    else:
                        match_sql = _expression_to_sql_for_order_by(
                            when_branch.get("match")
                        )
                        parts.append(f"WHEN {match_sql} THEN")
                    parts.append(
                        _expression_to_sql_for_order_by(when_branch.get("result"))
                    )
            if expression.get("else") is not None:
                parts.append("ELSE")
                parts.append(_expression_to_sql_for_order_by(expression["else"]))
            parts.append("END")
            return " ".join(parts)
    return str(expression)


def _annotate_column_tables(expression: Any) -> None:
    if not isinstance(expression, dict):
        return

    expression_type = expression.get("type")
    if expression_type == "column":
        source = expression.get("source")
        if isinstance(source, str):
            expression["table"] = source
        elif isinstance(expression.get("table"), str):
            expression["source"] = expression["table"]
        return

    if expression_type == "alias":
        _annotate_column_tables(expression.get("expression"))
        return

    if expression_type == "unary_op":
        _annotate_column_tables(expression.get("operand"))
        return

    if expression_type == "binary_op":
        _annotate_column_tables(expression.get("left"))
        _annotate_column_tables(expression.get("right"))
        return

    if expression_type == "function":
        args = expression.get("args")
        if isinstance(args, list):
            for argument in args:
                _annotate_column_tables(argument)
        return

    if expression_type == "cast":
        _annotate_column_tables(expression.get("value"))
        return

    if expression_type == "aggregate":
        filter_clause = expression.get("filter")
        if isinstance(filter_clause, dict):
            _annotate_column_tables(filter_clause)
        return

    if expression_type == "window_function":
        args = expression.get("args")
        if isinstance(args, list):
            for argument in args:
                _annotate_column_tables(argument)

        partition_by = expression.get("partition_by")
        if isinstance(partition_by, list):
            for partition_expression in partition_by:
                _annotate_column_tables(partition_expression)

        order_by = expression.get("order_by")
        if isinstance(order_by, list):
            for item in order_by:
                if not isinstance(item, dict):
                    continue
                order_expression = item.get("__expression__")
                if order_expression is not None:
                    _annotate_column_tables(order_expression)

        filter_clause = expression.get("filter")
        if isinstance(filter_clause, dict):
            _annotate_column_tables(filter_clause)
        return

    if expression_type == "case":
        _annotate_column_tables(expression.get("value"))
        for when_branch in expression.get("whens", []):
            if not isinstance(when_branch, dict):
                continue
            _annotate_column_tables(when_branch.get("match"))
            _annotate_column_tables(when_branch.get("condition"))
            _annotate_column_tables(when_branch.get("result"))
        _annotate_column_tables(expression.get("else"))
        return

    if "conditions" in expression:
        for condition in expression["conditions"]:
            _annotate_column_tables(condition)
        return

    if expression_type == "not":
        _annotate_column_tables(expression.get("operand"))
        return

    for key in ("column", "value"):
        _annotate_column_tables(expression.get(key))
