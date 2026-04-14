import re
from typing import Any, Dict, List, Optional, SupportsIndex, Union


class _QuotedString(str):
    pass


class _OrderByClause(list[dict[str, Any]]):
    def __getitem__(self, index: SupportsIndex | slice | str) -> Any:
        if isinstance(index, str):
            if len(self) != 1:
                raise TypeError("list indices must be integers or slices, not str")
            return super().__getitem__(0)[index]
        return super().__getitem__(index)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, dict):
            return len(self) == 1 and super().__getitem__(0) == other
        return super().__eq__(other)


def _is_placeholder(value: Any) -> bool:
    return value == "?" and not isinstance(value, _QuotedString)


def _split_csv(text: str) -> List[str]:
    items: List[str] = []
    current: List[str] = []
    in_single = False
    in_double = False
    paren_depth = 0
    case_depth = 0
    index = 0

    while index < len(text):
        char = text[index]

        if in_single:
            current.append(char)
            if char == "'":
                if index + 1 < len(text) and text[index + 1] == "'":
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_single = False
            index += 1
            continue

        if in_double:
            current.append(char)
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_double = False
            index += 1
            continue

        if char == "'":
            in_single = True
            current.append(char)
            index += 1
            continue

        if char == '"':
            in_double = True
            current.append(char)
            index += 1
            continue

        if char == "(":
            paren_depth += 1
            current.append(char)
            index += 1
            continue

        if char == ")":
            if paren_depth > 0:
                paren_depth -= 1
            current.append(char)
            index += 1
            continue

        if char.isalpha() or char == "_":
            start = index
            while index < len(text) and (text[index].isalnum() or text[index] == "_"):
                index += 1
            word = text[start:index]
            upper = word.upper()
            if upper == "CASE":
                case_depth += 1
            elif upper == "END" and case_depth > 0:
                case_depth -= 1
            current.append(word)
            continue

        if char == "," and paren_depth == 0 and case_depth == 0:
            items.append("".join(current).strip())
            current = []
            index += 1
            continue

        current.append(char)
        index += 1

    if current:
        items.append("".join(current).strip())
    return items


def _split_csv_preserve_empty(text: str) -> List[str]:
    items: List[str] = []
    current: List[str] = []
    in_single = False
    in_double = False
    paren_depth = 0
    case_depth = 0
    index = 0

    while index < len(text):
        char = text[index]

        if in_single:
            current.append(char)
            if char == "'":
                if index + 1 < len(text) and text[index + 1] == "'":
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_single = False
            index += 1
            continue

        if in_double:
            current.append(char)
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_double = False
            index += 1
            continue

        if char == "'":
            in_single = True
            current.append(char)
            index += 1
            continue

        if char == '"':
            in_double = True
            current.append(char)
            index += 1
            continue

        if char == "(":
            paren_depth += 1
            current.append(char)
            index += 1
            continue

        if char == ")":
            if paren_depth > 0:
                paren_depth -= 1
            current.append(char)
            index += 1
            continue

        if char.isalpha() or char == "_":
            start = index
            while index < len(text) and (text[index].isalnum() or text[index] == "_"):
                index += 1
            word = text[start:index]
            upper = word.upper()
            if upper == "CASE":
                case_depth += 1
            elif upper == "END" and case_depth > 0:
                case_depth -= 1
            current.append(word)
            continue

        if char == "," and paren_depth == 0 and case_depth == 0:
            items.append("".join(current).strip())
            current = []
            index += 1
            continue

        current.append(char)
        index += 1

    items.append("".join(current).strip())
    return items


_COLUMN_TYPE_ALIASES = {
    "INT": "INTEGER",
    "FLOAT": "REAL",
}

_SUPPORTED_COLUMN_TYPES = {"TEXT", "INTEGER", "REAL", "BOOLEAN", "DATE", "DATETIME"}


def _normalize_column_type(type_name: str, *, context: str) -> str:
    normalized = _COLUMN_TYPE_ALIASES.get(type_name.upper(), type_name.upper())
    if normalized not in _SUPPORTED_COLUMN_TYPES:
        raise ValueError(f"Unsupported {context} column type: {type_name}")
    return normalized


def _tokenize(text: str) -> List[str]:
    tokens: List[str] = []
    current: List[str] = []
    in_single = False
    in_double = False
    index = 0

    while index < len(text):
        char = text[index]

        if in_single:
            current.append(char)
            if char == "'":
                if index + 1 < len(text) and text[index + 1] == "'":
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_single = False
            index += 1
            continue

        if in_double:
            current.append(char)
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_double = False
            index += 1
            continue

        if char.isspace():
            if current:
                tokens.append("".join(current))
                current = []
            index += 1
            continue

        if char == "'":
            current.append(char)
            in_single = True
            index += 1
            continue

        if char == '"':
            current.append(char)
            in_double = True
            index += 1
            continue

        if char in {"(", ")"}:
            if current:
                tokens.append("".join(current))
                current = []
            tokens.append(char)
            index += 1
            continue

        current.append(char)
        index += 1

    if current:
        tokens.append("".join(current))

    return tokens


def _count_unquoted_placeholders(sql: str) -> int:
    """Count ``?`` placeholders outside string literals in *sql*."""
    count = 0
    in_quote = False
    quote_char = ""
    i = 0
    length = len(sql)
    while i < length:
        ch = sql[i]
        if in_quote:
            if ch == quote_char:
                if i + 1 < length and sql[i + 1] == quote_char:
                    i += 2
                    continue
                in_quote = False
        else:
            if ch in ("'", '"'):
                in_quote = True
                quote_char = ch
            elif ch == "?":
                count += 1
        i += 1
    return count


def _parse_value(token: str) -> Any:
    token = token.strip()
    if token.upper() == "NULL":
        return None
    if token.startswith("'") and token.endswith("'") and len(token) >= 2:
        # Unescape doubled single quotes: 'it''s' -> it's
        return _QuotedString(token[1:-1].replace("''", "'"))
    if token.startswith('"') and token.endswith('"') and len(token) >= 2:
        # Unescape doubled double quotes: "say ""hello""" -> say "hello"
        return _QuotedString(token[1:-1].replace('""', '"'))
    try:
        return int(token)
    except ValueError:
        pass
    try:
        return float(token)
    except ValueError:
        return token


def _parse_numeric_literal(token: str) -> int | float | None:
    if token.startswith(("'", '"')) and token.endswith(("'", '"')):
        return None
    try:
        return int(token)
    except ValueError:
        pass
    try:
        return float(token)
    except ValueError:
        return None


def _find_matching_parenthesis(tokens: List[str], start_index: int) -> int:
    if start_index >= len(tokens) or tokens[start_index] != "(":
        raise ValueError("Invalid SQL syntax: expected '('")

    depth = 0
    for index in range(start_index, len(tokens)):
        token = tokens[index]
        if token == "(":
            depth += 1
            continue
        if token == ")":
            depth -= 1
            if depth == 0:
                return index

    raise ValueError("Invalid SQL syntax: unmatched parenthesis")


def _find_top_level_keyword_index(tokens: List[str], keyword: str) -> int:
    depth = 0
    keyword_upper = keyword.upper()
    for index, token in enumerate(tokens):
        if token == "(":
            depth += 1
            continue
        if token == ")":
            if depth > 0:
                depth -= 1
            continue
        if depth == 0 and token.upper() == keyword_upper:
            return index
    return -1


def _tokenize_expression(text: str) -> list[str]:
    tokens: list[str] = []
    current: list[str] = []
    in_single = False
    in_double = False
    index = 0
    while index < len(text):
        char = text[index]

        if in_single:
            current.append(char)
            if char == "'":
                if index + 1 < len(text) and text[index + 1] == "'":
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_single = False
            index += 1
            continue

        if in_double:
            current.append(char)
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_double = False
            index += 1
            continue

        if char.isspace():
            if current:
                tokens.append("".join(current))
                current = []
            index += 1
            continue

        if char == "|" and index + 1 < len(text) and text[index + 1] == "|":
            if current:
                tokens.append("".join(current))
                current = []
            tokens.append("||")
            index += 2
            continue

        if char in {"+", "-", "*", "/", "(", ")", ","}:
            if current:
                tokens.append("".join(current))
                current = []
            tokens.append(char)
            index += 1
            continue

        if char == "'":
            current.append(char)
            in_single = True
            index += 1
            continue

        if char == '"':
            current.append(char)
            in_double = True
            index += 1
            continue

        current.append(char)
        index += 1

    if current:
        tokens.append("".join(current))

    return _collapse_aggregate_tokens(tokens)


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
    elif _is_quoted_token(expression):
        literal_value = _parse_value(expression)

    if literal_value is not None or expression.upper() == "NULL":
        return {"type": "literal", "value": literal_value}

    if re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, expression):
        source, name = expression.split(".", 1)
        return {"type": "column", "source": source, "name": name}

    if re.fullmatch(_IDENTIFIER_PATTERN, expression):
        return expression

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

        if token.upper() == "NULL" or token == "?" or _is_quoted_token(token):
            return {"type": "literal", "value": _parse_value(token)}

        if re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, token):
            source, name = token.split(".", 1)
            return {"type": "column", "source": source, "name": name}

        if re.fullmatch(_IDENTIFIER_PATTERN, token):
            return token

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


_AGGREGATE_FUNCTIONS = frozenset({"COUNT", "SUM", "AVG", "MIN", "MAX"})
_WINDOW_FUNCTIONS = frozenset({"ROW_NUMBER", "RANK", "DENSE_RANK"})
_SCALAR_FUNCTION_NAMES = frozenset(
    {
        "COALESCE",
        "NULLIF",
        "UPPER",
        "LOWER",
        "TRIM",
        "LENGTH",
        "SUBSTR",
        "SUBSTRING",
        "CONCAT",
        "YEAR",
        "MONTH",
        "DAY",
    }
)
_IDENTIFIER_PATTERN = r"[A-Za-z_][A-Za-z0-9_]*"
_QUALIFIED_IDENTIFIER_PATTERN = rf"{_IDENTIFIER_PATTERN}\.{_IDENTIFIER_PATTERN}"


def _aggregate_expression_to_label(aggregate: dict[str, Any]) -> str:
    func = str(aggregate.get("func", "")).upper()
    arg = str(aggregate.get("arg", "")).strip()
    aggregate_sql = (
        f"{func}(DISTINCT {arg})" if aggregate.get("distinct") else f"{func}({arg})"
    )
    filter_clause = aggregate.get("filter")
    if isinstance(filter_clause, dict):
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
        if not re.fullmatch(
            rf"{_IDENTIFIER_PATTERN}|{_QUALIFIED_IDENTIFIER_PATTERN}",
            arg,
        ):
            raise ValueError(
                f"Unsupported aggregate expression: COUNT(DISTINCT {arg}). "
                "Only bare and qualified column names are supported with DISTINCT"
            )

    if arg == "*" and func != "COUNT":
        raise ValueError(f"{func} does not support *")
    if arg != "*" and not re.fullmatch(
        rf"{_IDENTIFIER_PATTERN}|{_QUALIFIED_IDENTIFIER_PATTERN}",
        arg,
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


def _collapse_aggregate_tokens(tokens: List[str]) -> List[str]:
    collapsed: List[str] = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        upper = token.upper()
        if (
            upper == "COUNT"
            and index + 4 < len(tokens)
            and tokens[index + 1] == "("
            and tokens[index + 2].upper() == "DISTINCT"
            and tokens[index + 4] == ")"
        ):
            arg = tokens[index + 3].strip()
            collapsed.append(f"{upper}(DISTINCT {arg})")
            index += 5
            continue
        if (
            upper in _AGGREGATE_FUNCTIONS
            and index + 3 < len(tokens)
            and tokens[index + 1] == "("
            and tokens[index + 3] == ")"
        ):
            arg = tokens[index + 2].strip()
            collapsed.append(f"{upper}({arg})")
            index += 4
            continue
        collapsed.append(token)
        index += 1
    return collapsed


def _normalize_aggregate_expressions(text: str) -> str:
    return " ".join(_collapse_aggregate_tokens(_tokenize(text)))


def _is_quoted_token(token: str) -> bool:
    return len(token) >= 2 and (
        (token.startswith("'") and token.endswith("'"))
        or (token.startswith('"') and token.endswith('"'))
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
        if not re.fullmatch(_IDENTIFIER_PATTERN, alias_name):
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
            alias_name = candidate_alias
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
                    alias_name = candidate_alias

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


def _collect_qualified_references_from_expression(expression: Any) -> set[str]:
    refs: set[str] = set()
    if isinstance(expression, str):
        if re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, expression):
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
        if isinstance(arg, str) and re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, arg):
            refs.add(arg)
        filter_clause = expression.get("filter")
        if isinstance(filter_clause, dict):
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
                if isinstance(order_column, str) and re.fullmatch(
                    _QUALIFIED_IDENTIFIER_PATTERN,
                    order_column,
                ):
                    refs.add(order_column)

        filter_clause = expression.get("filter")
        if isinstance(filter_clause, dict):
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
        if re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, column):
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
                if re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, group_column):
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
            if isinstance(order_column, str) and re.fullmatch(
                _QUALIFIED_IDENTIFIER_PATTERN, order_column
            ):
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
        source = reference.split(".", 1)[0]
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
                raise ValueError(
                    "Aggregate functions are not allowed in WHERE clause; use HAVING instead"
                )
    if len(tokens) < 3:
        # Could be: NOT col = val (3+ tokens), (col = val) (5+ tokens)
        # Check for NOT with at least 2 following tokens
        if not (len(tokens) >= 1 and tokens[0].upper() == "NOT") and not (
            len(tokens) >= 1 and tokens[0] == "("
        ):
            raise ValueError("Invalid WHERE clause format")

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
        raise ValueError("Invalid WHERE clause format")

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
        raise ValueError("Invalid WHERE clause format")

    if tokens[index].upper() == "EXISTS":
        if not allow_subqueries:
            raise ValueError("Subqueries are not supported in this context")
        if index + 1 >= len(tokens) or tokens[index + 1] != "(":
            raise ValueError(
                "Invalid WHERE clause format: EXISTS requires '(SELECT ... )'"
            )

        subquery_end = _find_matching_parenthesis(tokens, index + 1)
        subquery_tokens = tokens[index + 2 : subquery_end]
        if not subquery_tokens or subquery_tokens[0].upper() != "SELECT":
            raise ValueError(
                "Invalid WHERE clause format: EXISTS requires SELECT subquery"
            )

        subquery_sql = " ".join(subquery_tokens).strip()
        if "?" in _tokenize(subquery_sql):
            raise ValueError(
                "Parameterized subqueries are not supported; use literal values"
            )

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
            raise ValueError("Invalid WHERE clause format: unmatched parenthesis")
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
        raise ValueError("Invalid WHERE clause format")

    if _is_case_keyword(expression_tokens[0], "CASE"):
        parsed_case, consumed = _parse_case_expression_tokens(expression_tokens, 0)
        if consumed != len(expression_tokens):
            raise ValueError("Invalid WHERE clause format")
        return parsed_case

    if (
        expression_tokens[0] == "("
        and len(expression_tokens) > 2
        and expression_tokens[1].upper() == "SELECT"
    ):
        if not allow_subqueries:
            raise ValueError("Subqueries are not supported in this context")
        subquery_end = _find_matching_parenthesis(expression_tokens, 0)
        if subquery_end != len(expression_tokens) - 1:
            raise ValueError("Invalid WHERE clause format")
        subquery_sql = " ".join(expression_tokens[1:subquery_end]).strip()
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
    except ValueError as exc:
        if (
            allow_aggregates
            and "Unsupported aggregate expression" in str(exc)
            and re.fullmatch(r"(?i)(COUNT|SUM|AVG|MIN|MAX)\s*\(.+\)", expression_text)
        ):
            return _normalize_aggregate_expressions(expression_text)
        raise ValueError("Invalid WHERE clause format") from exc

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
        raise ValueError("Invalid WHERE clause format")

    expression_tokens, expression_end = _collect_condition_expression_tokens(
        tokens,
        index,
        stop_at_operator=True,
    )
    if not expression_tokens:
        raise ValueError("Invalid WHERE clause format")

    try:
        parsed_operand = _parse_condition_expression_tokens(
            expression_tokens,
            allow_aggregates=allow_aggregates,
            allow_subqueries=allow_subqueries,
            outer_sources=outer_sources,
            collapse_literals=False,
        )
    except ValueError as exc:
        if "Unsupported column expression" in str(exc):
            raise ValueError("Invalid WHERE clause format") from exc
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
        raise ValueError("Invalid WHERE clause format")

    value_stop_keywords = stop_keywords or {"AND", "OR"}
    expression_tokens, expression_end = _collect_condition_expression_tokens(
        tokens,
        index,
        stop_keywords=value_stop_keywords,
    )
    if not expression_tokens:
        raise ValueError("Invalid WHERE clause format")

    try:
        parsed_value = _parse_condition_expression_tokens(
            expression_tokens,
            allow_aggregates=allow_aggregates,
            allow_subqueries=allow_subqueries,
            outer_sources=outer_sources,
            collapse_literals=True,
        )
    except ValueError as exc:
        if "Unsupported column expression" in str(exc):
            raise ValueError("Invalid WHERE clause format") from exc
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
        raise ValueError("Invalid WHERE clause format")

    column, operator_index = _parse_condition_operand(
        tokens,
        index,
        allow_subqueries=allow_subqueries,
        allow_aggregates=allow_aggregates,
        outer_sources=outer_sources,
    )
    if operator_index >= len(tokens):
        raise ValueError("Invalid WHERE clause format")

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
            raise ValueError("Invalid WHERE clause format: expected NULL after IS NOT")
        if (
            operator_index + 1 < len(tokens)
            and tokens[operator_index + 1].upper() == "NULL"
        ):
            return (
                {"column": column, "operator": "IS", "value": None},
                operator_index + 2,
            )
        raise ValueError("Invalid WHERE clause format: expected NULL or NOT after IS")

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
        raise ValueError(
            "Invalid WHERE clause format: expected IN, LIKE, ILIKE, or BETWEEN after NOT"
        )

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
        raise ValueError("Invalid WHERE clause format")
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
        raise ValueError("Invalid WHERE clause format: expected AND in BETWEEN clause")

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
        raise ValueError("Invalid WHERE clause format")

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
            raise ValueError(
                "Invalid WHERE clause format: ESCAPE requires a single character"
            )
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
        raise ValueError("Invalid WHERE clause format")

    in_start = paren_start
    in_end = in_start
    paren_depth = 0
    while in_end < len(tokens):
        token = tokens[in_end]
        if token == "(":
            paren_depth += 1
        elif token == ")":
            if paren_depth == 0:
                raise ValueError("Invalid WHERE clause format: malformed IN clause")
            paren_depth -= 1
            if paren_depth == 0:
                break
        in_end += 1
    if in_end >= len(tokens):
        raise ValueError("Invalid WHERE clause format: expected ')' in IN clause")

    in_values_text = " ".join(tokens[in_start : in_end + 1]).strip()
    if not in_values_text.startswith("(") or not in_values_text.endswith(")"):
        raise ValueError("Invalid WHERE clause format: malformed IN clause")

    raw_values = in_values_text[1:-1].strip()
    if not raw_values:
        raise ValueError("Invalid WHERE clause format: IN clause cannot be empty")

    op = "NOT IN" if negated else "IN"

    raw_tokens = raw_values.split()
    if raw_tokens and raw_tokens[0].upper() == "SELECT":
        if not allow_subqueries:
            raise ValueError("Subqueries are not supported in this context")

        if "?" in _tokenize(raw_values):
            raise ValueError(
                "Parameterized subqueries are not supported; use literal values"
            )

        # Reject JOINs inside subqueries before other checks
        subquery_tokens = _tokenize(raw_values)
        for raw_token in subquery_tokens:
            if raw_token.startswith("'") or raw_token.startswith('"'):
                continue
            if raw_token.upper() == "JOIN":
                raise ValueError("JOIN is not supported in subqueries")

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
            raise ValueError("Correlated subqueries are not supported")

        if subquery_parsed.get("having") is not None:
            raise ValueError("HAVING is not supported in subqueries")
        if subquery_parsed.get("order_by") is not None:
            raise ValueError("ORDER BY is not supported in subqueries")
        if subquery_parsed.get("offset") is not None:
            raise ValueError("OFFSET is not supported in subqueries")
        if subquery_parsed.get("group_by") is not None:
            raise ValueError("GROUP BY is not supported in subqueries")
        if subquery_parsed.get("limit") is not None:
            raise ValueError("LIMIT is not supported in subqueries")
        if subquery_parsed.get("joins"):
            raise ValueError("JOIN is not supported in subqueries")

        subquery_columns = subquery_parsed["columns"]
        if subquery_columns == ["*"] or len(subquery_columns) != 1:
            raise ValueError("Subquery in WHERE ... IN must select exactly one column")

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
            raise ValueError("Invalid WHERE clause format: IN clause cannot be empty")

        return (
            {
                "column": column,
                "operator": op,
                "value": parsed_values,
            },
            in_end + 1,
        )


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
    if not re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, token):
        raise ValueError(
            f"Invalid column reference in JOIN clause: {token}. Expected source.column"
        )
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
        if not re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, column):
            raise ValueError(
                f"{context} requires qualified column names in JOIN queries"
            )
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
                if not re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, arg):
                    raise ValueError(
                        "Aggregate arguments in JOIN queries must be qualified column names or *"
                    )
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
                source, _ = reference.split(".", 1)
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


def _expression_to_sql_for_order_by(expression: Any) -> str:
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
        and re.fullmatch(
            rf"(?:{_IDENTIFIER_PATTERN}|{_QUALIFIED_IDENTIFIER_PATTERN})",
            expression_tokens[0],
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

    if re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, expression_text):
        return {"column": expression_text, "direction": direction}
    if re.fullmatch(_IDENTIFIER_PATTERN, expression_text):
        return {"column": expression_text, "direction": direction}

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
    table = tokens[from_index + 1]
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
            from_entry["alias"] = maybe_alias
            from_entry["ref"] = maybe_alias
            token_index += 1

    joins: List[Dict[str, Any]] = []
    known_source_refs: set[str] = {str(from_entry["table"]), str(from_entry["ref"])}
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
        join_table = tokens[token_index]
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
                join_source["alias"] = maybe_alias
                join_source["ref"] = maybe_alias
                token_index += 1

        right_sources = {str(join_source["table"]), str(join_source["ref"])}
        collision = known_source_refs & right_sources
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

                if not re.fullmatch(
                    _QUALIFIED_IDENTIFIER_PATTERN, qualified_group_column
                ):
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
                    if arg != "*" and not re.fullmatch(
                        _QUALIFIED_IDENTIFIER_PATTERN, arg
                    ):
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
                    if arg != "*" and not re.fullmatch(
                        _QUALIFIED_IDENTIFIER_PATTERN,
                        arg,
                    ):
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


def _split_insert_on_conflict_clause(remainder: str) -> tuple[str, Optional[str]]:
    tokens = _tokenize(remainder.strip())
    if not tokens:
        return "", None

    depth = 0
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if token == "(":
            depth += 1
        elif token == ")":
            if depth > 0:
                depth -= 1
        elif (
            depth == 0
            and token.upper() == "ON"
            and index + 1 < len(tokens)
            and tokens[index + 1].upper() == "CONFLICT"
        ):
            before = " ".join(tokens[:index]).strip()
            on_conflict = " ".join(tokens[index:]).strip()
            return before, on_conflict
        index += 1

    return remainder.strip(), None


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


def _parse_upsert_assignment_value(value_text: str) -> Any:
    stripped = value_text.strip()
    if not stripped:
        raise ValueError("Invalid ON CONFLICT clause format")

    is_expression = bool(re.match(r"(?i)^CASE\b", stripped)) or bool(
        re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, stripped)
    )
    if not is_expression:
        expression_tokens = _tokenize_expression(stripped)
        is_expression = any(
            token in {"+", "-", "*", "/", "||", "(", ")"} for token in expression_tokens
        )

    if is_expression:
        parsed = _parse_column_expression(
            stripped,
            allow_wildcard=False,
            allow_aggregates=False,
        )
        _annotate_column_tables(parsed)
        return parsed

    return _parse_value(stripped)


def _parse_on_conflict_clause(
    clause: str,
    query: str,
    params: Optional[tuple[Any, ...]],
) -> Dict[str, Any]:
    tokens = _tokenize(clause.strip())
    if len(tokens) < 5:
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
    if tokens[0].upper() != "ON" or tokens[1].upper() != "CONFLICT":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    index = 2
    if index >= len(tokens) or tokens[index] != "(":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    index += 1
    depth = 1
    target_tokens: List[str] = []
    while index < len(tokens) and depth > 0:
        token = tokens[index]
        if token == "(":
            depth += 1
            target_tokens.append(token)
        elif token == ")":
            depth -= 1
            if depth > 0:
                target_tokens.append(token)
        else:
            target_tokens.append(token)
        index += 1

    if depth != 0:
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    target_columns = _parse_columns(" ".join(target_tokens))
    invalid_target_columns = [
        col for col in target_columns if not isinstance(col, str) or col == "*"
    ]
    if invalid_target_columns:
        raise ValueError("ON CONFLICT target supports only bare column names")

    if index >= len(tokens) or tokens[index].upper() != "DO":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
    index += 1

    if index >= len(tokens):
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    action = tokens[index].upper()
    index += 1

    if action == "NOTHING":
        if index != len(tokens):
            raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
        if params is not None:
            _bind_params([], params)
        return {
            "target_columns": target_columns,
            "action": "NOTHING",
        }

    if action != "UPDATE":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    if index >= len(tokens) or tokens[index].upper() != "SET":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
    index += 1

    set_part = " ".join(tokens[index:]).strip()
    if not set_part:
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    assignments = []
    raw_assignments = _split_csv(set_part)
    for assignment in raw_assignments:
        if "=" not in assignment:
            raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
        col, value = assignment.split("=", 1)
        assignments.append(
            {
                "column": col.strip(),
                "value": _parse_upsert_assignment_value(value),
            }
        )

    values_to_bind: List[Any] = []
    for item in assignments:
        assignment_value = item["value"]
        if isinstance(assignment_value, dict):
            values_to_bind.extend(_expression_values_to_bind(assignment_value))
        else:
            values_to_bind.append(assignment_value)

    if params is not None or any(_is_placeholder(value) for value in values_to_bind):
        bound = _bind_params(values_to_bind, params)
        consumed = 0
        for item in assignments:
            assignment_value = item["value"]
            if isinstance(assignment_value, dict):
                consumed += _bind_expression_values(assignment_value, bound, consumed)
            else:
                item["value"] = bound[consumed]
                consumed += 1

    return {
        "target_columns": target_columns,
        "action": "UPDATE",
        "set": assignments,
    }


def _parse_insert(query: str, params: Optional[tuple[Any, ...]]) -> Dict[str, Any]:
    stripped = query.strip()
    upper = stripped.upper()
    insert_prefix = "INSERT INTO "
    if not upper.startswith(insert_prefix):
        raise ValueError(f"Invalid INSERT format: {query}")

    remainder = stripped[len(insert_prefix) :].strip()
    if not remainder:
        raise ValueError(f"Invalid INSERT format: {query}")

    split_index = 0
    while (
        split_index < len(remainder)
        and not remainder[split_index].isspace()
        and remainder[split_index] != "("
    ):
        split_index += 1
    table = remainder[:split_index].strip()
    if not table:
        raise ValueError(f"Invalid INSERT format: {query}")
    remainder = remainder[split_index:].strip()

    columns = None
    if remainder.startswith("("):
        depth = 0
        close_index = -1
        for idx, char in enumerate(remainder):
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    close_index = idx
                    break
            if depth < 0:
                break
        if close_index < 0:
            raise ValueError(f"Invalid INSERT format: {query}")
        cols_part = remainder[1:close_index]
        columns = _parse_columns(cols_part)
        invalid_columns = [
            col for col in columns if not isinstance(col, str) or col == "*"
        ]
        if invalid_columns:
            raise ValueError("INSERT column list supports only bare column names")
        remainder = remainder[close_index + 1 :].strip()

    if not remainder:
        raise ValueError(f"Invalid INSERT format: {query}")

    source_part, on_conflict_clause = _split_insert_on_conflict_clause(remainder)
    if not source_part:
        raise ValueError(f"Invalid INSERT format: {query}")

    remainder_upper = source_part.upper()
    remaining_params: Optional[tuple[Any, ...]] = None
    if remainder_upper.startswith("SELECT"):
        select_params = params
        if params is not None and on_conflict_clause is not None:
            select_param_count = _count_unquoted_placeholders(source_part)
            select_params = params[:select_param_count]
            remaining_params = params[select_param_count:]
        subquery = _parse_select(source_part, select_params)
        if remaining_params is None:
            remaining_params = () if params is not None else None
        values: Any = {"type": "subquery", "query": subquery}
    elif remainder_upper.startswith("VALUES"):
        values_part = source_part[len("VALUES") :].strip()
        if not values_part:
            raise ValueError(f"Invalid INSERT format: {query}")

        raw_rows: List[str] = []
        depth = 0
        row_start = -1
        expect_separator = False
        in_single = False
        in_double = False
        index = 0
        while index < len(values_part):
            char = values_part[index]
            if in_single:
                if char == "'":
                    if index + 1 < len(values_part) and values_part[index + 1] == "'":
                        index += 1
                    else:
                        in_single = False
                index += 1
                continue
            if in_double:
                if char == '"':
                    if index + 1 < len(values_part) and values_part[index + 1] == '"':
                        index += 1
                    else:
                        in_double = False
                index += 1
                continue

            if char == "'":
                in_single = True
                index += 1
                continue
            if char == '"':
                in_double = True
                index += 1
                continue

            if char.isspace() and depth == 0:
                index += 1
                continue

            if depth == 0:
                if expect_separator:
                    if char == ",":
                        expect_separator = False
                        index += 1
                        continue
                    raise ValueError(f"Invalid INSERT format: {query}")
                if char != "(":
                    raise ValueError(f"Invalid INSERT format: {query}")
                row_start = index + 1
                depth = 1
                index += 1
                continue

            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    if row_start < 0:
                        raise ValueError(f"Invalid INSERT format: {query}")
                    raw_rows.append(values_part[row_start:index])
                    expect_separator = True

            if depth < 0:
                raise ValueError(f"Invalid INSERT format: {query}")
            index += 1

        if depth != 0 or in_single or in_double or not raw_rows or not expect_separator:
            raise ValueError(f"Invalid INSERT format: {query}")

        if params is None:
            values = [
                _bind_params(
                    [_parse_value(token) for token in _split_csv(raw_row)], None
                )
                for raw_row in raw_rows
            ]
        else:
            bound_rows: List[List[Any]] = []
            param_index = 0
            for raw_row in raw_rows:
                parsed_row = [_parse_value(token) for token in _split_csv(raw_row)]
                placeholders = [value for value in parsed_row if _is_placeholder(value)]
                needed = len(placeholders)
                row_params = params[param_index : param_index + needed]
                bound_rows.append(_bind_params(parsed_row, row_params))
                param_index += needed
            remaining_params = params[param_index:]
            values = bound_rows
    else:
        raise ValueError(f"Invalid INSERT format: {query}")

    on_conflict = None
    if on_conflict_clause is not None:
        on_conflict = _parse_on_conflict_clause(
            on_conflict_clause, query, remaining_params
        )
        remaining_params = () if params is not None else None

    if (
        params is not None
        and remaining_params is not None
        and len(remaining_params) > 0
    ):
        raise ValueError("Too many parameters for placeholders")

    result: Dict[str, Any] = {
        "action": "INSERT",
        "table": table,
        "columns": columns,
        "values": values,
    }
    if on_conflict is not None:
        result["on_conflict"] = on_conflict
    return result


def _parse_create(query: str) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 3 or tokens[0].upper() != "CREATE" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    table_and_cols = " ".join(tokens[2:]).strip()
    if "(" not in table_and_cols or not table_and_cols.endswith(")"):
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    table_name, cols_part = table_and_cols.split("(", 1)
    table = table_name.strip()
    if not table:
        raise ValueError("Table name is required in CREATE TABLE")
    cols_part = cols_part.rsplit(")", 1)[0]
    raw_columns = _split_csv_preserve_empty(cols_part)
    empty_indexes = [
        index for index, definition in enumerate(raw_columns) if not definition.strip()
    ]
    has_single_trailing_empty = (
        len(empty_indexes) == 1 and empty_indexes[0] == len(raw_columns) - 1
    )
    if empty_indexes and not has_single_trailing_empty:
        raise ValueError("Malformed column definitions: empty column definition found")
    columns = []
    column_definitions = []
    for col in raw_columns:
        if not col.strip():
            continue
        stripped_col = col.strip()
        parts = stripped_col.split()
        if len(parts) > 2:
            raise ValueError(
                "Malformed column definition: "
                f"{stripped_col!r}. Missing comma between column definitions?"
            )
        column_name = parts[0]
        type_name = "TEXT"
        if len(parts) == 2:
            type_name = _normalize_column_type(parts[1], context="CREATE TABLE")
        columns.append(column_name)
        column_definitions.append({"name": column_name, "type_name": type_name})
    if not columns:
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    return {
        "action": "CREATE",
        "table": table,
        "columns": columns,
        "column_definitions": column_definitions,
    }


def _parse_drop(query: str) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) != 3 or tokens[0].upper() != "DROP" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid DROP TABLE format: {query}")
    return {
        "action": "DROP",
        "table": tokens[2],
    }


def _parse_alter(query: str) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 6 or tokens[0].upper() != "ALTER" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid ALTER TABLE format: {query}")

    table = tokens[2]
    operation = tokens[3].upper()

    if operation == "ADD":
        if len(tokens) != 7 or tokens[4].upper() != "COLUMN":
            raise ValueError(f"Invalid ALTER TABLE format: {query}")
        type_name = _normalize_column_type(tokens[6], context="ALTER TABLE")
        return {
            "action": "ALTER",
            "table": table,
            "operation": "ADD_COLUMN",
            "column": tokens[5],
            "type_name": type_name,
        }

    if operation == "DROP":
        if len(tokens) != 6 or tokens[4].upper() != "COLUMN":
            raise ValueError(f"Invalid ALTER TABLE format: {query}")
        return {
            "action": "ALTER",
            "table": table,
            "operation": "DROP_COLUMN",
            "column": tokens[5],
        }

    if operation == "RENAME":
        if (
            len(tokens) != 8
            or tokens[4].upper() != "COLUMN"
            or tokens[6].upper() != "TO"
        ):
            raise ValueError(f"Invalid ALTER TABLE format: {query}")
        return {
            "action": "ALTER",
            "table": table,
            "operation": "RENAME_COLUMN",
            "old_column": tokens[5],
            "new_column": tokens[7],
        }

    raise ValueError(f"Invalid ALTER TABLE format: {query}")


def _strip_outer_parens(tokens: List[str]) -> List[str]:
    """Strip one layer of outer parentheses from a token list.

    If the token list is wrapped in balanced ``(`` / ``)`` brackets that
    enclose the entire list, return the inner tokens.  Otherwise return
    the list unchanged.
    """
    if len(tokens) < 2 or tokens[0] != "(" or tokens[-1] != ")":
        return tokens
    depth = 0
    for i, tok in enumerate(tokens):
        if tok == "(":
            depth += 1
        elif tok == ")":
            depth -= 1
            if depth == 0 and i < len(tokens) - 1:
                # Closing paren is not the last token, so the outer parens
                # do NOT wrap the whole list.
                return tokens
    return tokens[1:-1]


def _extract_trailing_clauses(
    tokens: List[str],
) -> tuple[
    List[str],
    list[dict[str, Any]] | None,
    Optional[Union[int, str]],
    Optional[Union[int, str]],
]:
    """Split trailing ORDER BY / LIMIT / OFFSET from a compound's last branch.

    Returns ``(branch_tokens, order_by, limit, offset)``.
    """
    order_by: list[dict[str, Any]] | None = None
    limit: Optional[Union[int, str]] = None
    offset: Optional[Union[int, str]] = None

    # Scan backwards (at depth 0) for ORDER BY / LIMIT / OFFSET.
    # We work on the raw token list and only look for top-level keywords.
    uppers = [t.upper() for t in tokens]
    depth = 0
    keyword_positions: Dict[str, int] = {}
    for i, tok in enumerate(tokens):
        if tok == "(":
            depth += 1
        elif tok == ")":
            depth -= 1
        elif depth == 0:
            u = uppers[i]
            if u == "OFFSET" and "OFFSET" not in keyword_positions:
                keyword_positions["OFFSET"] = i
            elif u == "LIMIT" and "LIMIT" not in keyword_positions:
                keyword_positions["LIMIT"] = i
            elif u == "ORDER" and i + 1 < len(uppers) and uppers[i + 1] == "BY":
                if "ORDER BY" not in keyword_positions:
                    keyword_positions["ORDER BY"] = i

    if not keyword_positions:
        return tokens, order_by, limit, offset

    # Determine the cut point (earliest trailing clause).
    cut = min(keyword_positions.values())

    # Only accept trailing clauses that come AFTER the FROM clause of the
    # last branch. A simple heuristic: the cut must come after the last
    # top-level FROM token.
    last_from = -1
    from_depth = 0
    for i, u in enumerate(uppers[:cut]):
        tok = tokens[i]
        if tok == "(":
            from_depth += 1
        elif tok == ")":
            from_depth -= 1
        elif u == "FROM" and from_depth == 0:
            last_from = i
    if last_from == -1:
        return tokens, order_by, limit, offset

    branch_tokens = tokens[:cut]
    trail_tokens = tokens[cut:]
    trail_uppers = [t.upper() for t in trail_tokens]

    idx = 0
    while idx < len(trail_tokens):
        tu = trail_uppers[idx]
        if (
            tu == "ORDER"
            and idx + 1 < len(trail_uppers)
            and trail_uppers[idx + 1] == "BY"
        ):
            idx += 2  # skip ORDER BY
            order_tokens_list: list[str] = []
            while idx < len(trail_tokens) and trail_uppers[idx] not in {
                "LIMIT",
                "OFFSET",
            }:
                order_tokens_list.append(trail_tokens[idx])
                idx += 1
            order_by = _parse_order_by_clause_tokens(order_tokens_list)
        elif tu == "LIMIT":
            idx += 1
            if idx >= len(trail_tokens):
                raise ValueError("Invalid LIMIT clause format")
            val = trail_tokens[idx]
            limit = int(val) if val != "?" else val
            idx += 1
        elif tu == "OFFSET":
            idx += 1
            if idx >= len(trail_tokens):
                raise ValueError("Invalid OFFSET clause format")
            val = trail_tokens[idx]
            offset = int(val) if val != "?" else val
            idx += 1
        else:
            # Unknown trailing token — not a compound-level clause, put it back.
            return tokens, None, None, None

    return branch_tokens, order_by, limit, offset


def _parse_compound(
    query: str, params: Optional[tuple[Any, ...]]
) -> Optional[Dict[str, Any]]:
    tokens = _tokenize(query.strip())
    if not tokens:
        return None

    # Accept queries starting with either SELECT or ( (parenthesized branch).
    first_upper = tokens[0].upper()
    if first_upper != "SELECT" and tokens[0] != "(":
        return None

    query_tokens: List[List[str]] = []
    operators: List[str] = []
    current_tokens: List[str] = []
    depth = 0
    index = 0
    found_compound = False

    while index < len(tokens):
        token = tokens[index]
        upper = token.upper()

        if token == "(":
            depth += 1
            current_tokens.append(token)
            index += 1
            continue
        if token == ")":
            depth -= 1
            if depth < 0:
                if not found_compound:
                    return None
                raise ValueError(f"Invalid SQL query format: {query}")
            current_tokens.append(token)
            index += 1
            continue

        if depth == 0 and upper in {"UNION", "INTERSECT", "EXCEPT"}:
            if not current_tokens:
                raise ValueError(f"Invalid SQL query format: {query}")
            query_tokens.append(current_tokens)
            current_tokens = []
            found_compound = True

            if (
                upper == "UNION"
                and index + 1 < len(tokens)
                and tokens[index + 1].upper() == "ALL"
            ):
                operators.append("UNION ALL")
                index += 2
            else:
                operators.append(upper)
                index += 1
            continue

        current_tokens.append(token)
        index += 1

    if depth != 0 and found_compound:
        raise ValueError(f"Invalid SQL query format: {query}")
    if not found_compound:
        return None
    if not current_tokens:
        raise ValueError(f"Invalid SQL query format: {query}")

    query_tokens.append(current_tokens)
    if len(query_tokens) != len(operators) + 1:
        raise ValueError(f"Invalid SQL query format: {query}")

    # Extract compound-level ORDER BY / LIMIT / OFFSET from the last branch.
    last_branch = query_tokens[-1]
    last_branch, compound_order_by, compound_limit, compound_offset = (
        _extract_trailing_clauses(last_branch)
    )
    query_tokens[-1] = last_branch

    # Strip outer parens from each branch (e.g. "(SELECT ... ORDER BY ...)")
    query_tokens = [_strip_outer_parens(branch) for branch in query_tokens]

    parsed_queries: List[Dict[str, Any]] = []
    param_index = 0
    total_params = len(params) if params is not None else 0
    for segment_tokens in query_tokens:
        if not segment_tokens or segment_tokens[0].upper() != "SELECT":
            raise ValueError("Compound queries support only SELECT subqueries")

        segment_query = " ".join(segment_tokens)
        segment_params: Optional[tuple[Any, ...]] = None
        if params is not None:
            placeholder_count = _count_unquoted_placeholders(segment_query)
            next_index = param_index + placeholder_count
            if next_index > total_params:
                raise ValueError("Not enough parameters for placeholders")
            segment_params = params[param_index:next_index]
            param_index = next_index

        parsed_queries.append(_parse_select(segment_query, segment_params))

    # Resolve compound-level LIMIT / OFFSET placeholders ("?").
    if params is not None:
        if compound_limit == "?":
            if param_index >= total_params:
                raise ValueError("Not enough parameters for LIMIT placeholder")
            compound_limit = int(params[param_index])
            param_index += 1
        if compound_offset == "?":
            if param_index >= total_params:
                raise ValueError("Not enough parameters for OFFSET placeholder")
            compound_offset = int(params[param_index])
            param_index += 1

    if params is not None and param_index < total_params:
        raise ValueError("Too many parameters for placeholders")

    result: Dict[str, Any] = {
        "action": "COMPOUND",
        "operator": operators[0],
        "operators": operators,
        "queries": parsed_queries,
    }
    if compound_order_by is not None:
        result["order_by"] = compound_order_by
    if compound_limit is not None:
        result["limit"] = compound_limit
    if compound_offset is not None:
        result["offset"] = compound_offset
    return result


def _query_references_name(query: Dict[str, Any], name: str) -> bool:
    lowered = name.lower()
    action = str(query.get("action", "")).upper()
    if action == "COMPOUND":
        for branch in query.get("queries", []):
            if isinstance(branch, dict) and _query_references_name(branch, name):
                return True
        return False

    if action != "SELECT":
        return False

    for source in _query_source_references(query):
        if source.lower() == lowered:
            return True
    return False


def _parse_with_query(
    query: str,
    params: Optional[tuple[Any, ...]],
) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 2 or tokens[0].upper() != "WITH":
        raise ValueError(f"Invalid SQL query format: {query}")

    index = 1
    if tokens[index].upper() == "RECURSIVE":
        raise ValueError("Recursive CTEs are not supported")

    ctes: List[Dict[str, Any]] = []
    cte_names: set[str] = set()
    param_index = 0
    total_params = len(params) if params is not None else 0

    while True:
        if index >= len(tokens):
            raise ValueError("Invalid WITH clause: missing CTE name")

        cte_name = tokens[index]
        if not re.fullmatch(_IDENTIFIER_PATTERN, cte_name):
            raise ValueError(f"Invalid CTE name: {cte_name}")
        lowered_name = cte_name.lower()
        if lowered_name in cte_names:
            raise ValueError(f"Duplicate CTE name: {cte_name}")
        cte_names.add(lowered_name)
        index += 1

        if index >= len(tokens) or tokens[index].upper() != "AS":
            raise ValueError("Invalid WITH clause: expected AS")
        index += 1

        if index >= len(tokens) or tokens[index] != "(":
            raise ValueError("Invalid WITH clause: expected '(' after AS")
        end_index = _find_matching_parenthesis(tokens, index)
        subquery_tokens = tokens[index + 1 : end_index]
        if not subquery_tokens:
            raise ValueError("Invalid WITH clause: empty CTE query")

        subquery_sql = " ".join(subquery_tokens)
        cte_params: Optional[tuple[Any, ...]] = None
        if params is not None:
            placeholder_count = _count_unquoted_placeholders(subquery_sql)
            next_param_index = param_index + placeholder_count
            if next_param_index > total_params:
                raise ValueError("Not enough parameters for placeholders")
            cte_params = params[param_index:next_param_index]
            param_index = next_param_index

        parsed_cte = _parse_compound(subquery_sql, cte_params)
        if parsed_cte is None:
            parsed_cte = _parse_select(subquery_sql, cte_params)

        if _query_references_name(parsed_cte, cte_name):
            raise ValueError("Recursive CTEs are not supported")

        ctes.append({"type": "cte", "name": cte_name, "query": parsed_cte})
        index = end_index + 1

        if index >= len(tokens):
            raise ValueError("Invalid WITH clause: missing main SELECT query")
        if tokens[index] == ",":
            index += 1
            continue
        break

    main_query = " ".join(tokens[index:]).strip()
    if not main_query:
        raise ValueError("Invalid WITH clause: missing main SELECT query")

    main_tokens = _tokenize(main_query)
    if not main_tokens:
        raise ValueError("Invalid WITH clause: missing main SELECT query")
    if main_tokens[0].upper() != "SELECT" and main_tokens[0] != "(":
        raise ValueError("WITH clause requires a SELECT query")

    remaining_params: Optional[tuple[Any, ...]] = None
    if params is not None:
        remaining_params = params[param_index:]

    parsed = _parse_compound(main_query, remaining_params)
    if parsed is None:
        if main_tokens[0] == "(":
            raise ValueError(f"Unsupported SQL action: {main_tokens[0]}")
        parsed = _parse_select(main_query, remaining_params)

    parsed["ctes"] = ctes
    return parsed


def parse_sql(query: str, params: Optional[tuple[Any, ...]] = None) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if not tokens:
        raise ValueError(f"Invalid SQL query format: {query}")
    action = tokens[0].upper()
    parsed: Dict[str, Any]
    if action == "WITH":
        parsed = _parse_with_query(query, params)
    elif action == "SELECT" or tokens[0] == "(":
        compound_parsed = _parse_compound(query, params)
        if compound_parsed is None:
            if tokens[0] == "(":
                raise ValueError(f"Unsupported SQL action: {tokens[0]}")
            parsed = _parse_select(query, params)
        else:
            parsed = compound_parsed
    elif action == "INSERT":
        parsed = _parse_insert(query, params)
    elif action == "CREATE":
        parsed = _parse_create(query)
    elif action == "DROP":
        parsed = _parse_drop(query)
    elif action == "UPDATE":
        parsed = _parse_update(query, params)
    elif action == "DELETE":
        parsed = _parse_delete(query, params)
    elif action == "ALTER":
        parsed = _parse_alter(query)
    else:
        raise ValueError(f"Unsupported SQL action: {action}")

    parsed["params"] = params
    return parsed


def _parse_update(query: str, params: Optional[tuple[Any, ...]]) -> Dict[str, Any]:
    upper = query.upper()
    if " SET " not in upper:
        raise ValueError(f"Invalid UPDATE format: {query}")
    set_index = upper.index(" SET ")
    before_set = query[:set_index]
    after_set = query[set_index + len(" SET ") :]
    before_tokens = _tokenize(before_set.strip())
    if len(before_tokens) < 2 or before_tokens[0].upper() != "UPDATE":
        raise ValueError(f"Invalid UPDATE format: {query}")
    table = before_tokens[1].strip()

    where_part = None
    after_upper = after_set.upper()
    if " WHERE " in after_upper:
        where_index = after_upper.index(" WHERE ")
        set_part = after_set[:where_index]
        where_part = after_set[where_index + len(" WHERE ") :]
    else:
        set_part = after_set

    assignments = []
    raw_assignments = _split_csv(set_part.strip())
    for assignment in raw_assignments:
        if "=" not in assignment:
            raise ValueError(f"Invalid UPDATE format: {query}")
        col, value = assignment.split("=", 1)
        parsed_value: Any
        value_text = value.strip()
        if re.match(r"(?i)^CASE\b", value_text):
            parsed_value = _parse_column_expression(
                value_text,
                allow_wildcard=False,
                allow_aggregates=False,
            )
        else:
            parsed_value = _parse_value(value)
        assignments.append({"column": col.strip(), "value": parsed_value})

    where = None
    if where_part:
        where = _parse_where_expression(
            where_part,
            params,
            bind_params=False,
            allow_subqueries=True,
            outer_sources={table},
        )

    values_to_bind: List[Any] = []
    for item in assignments:
        assignment_value = item["value"]
        if isinstance(assignment_value, dict):
            values_to_bind.extend(_expression_values_to_bind(assignment_value))
        else:
            values_to_bind.append(assignment_value)
    if where is not None:
        values_to_bind.extend(_where_values_to_bind(where))
    if params is not None or any(_is_placeholder(value) for value in values_to_bind):
        bound = _bind_params(values_to_bind, params)
        consumed = 0
        for item in assignments:
            assignment_value = item["value"]
            if isinstance(assignment_value, dict):
                consumed += _bind_expression_values(assignment_value, bound, consumed)
            else:
                item["value"] = bound[consumed]
                consumed += 1
        if where is not None:
            _bind_where_conditions(where, bound, consumed)

    return {
        "action": "UPDATE",
        "table": table,
        "set": assignments,
        "where": where,
    }


def _parse_delete(query: str, params: Optional[tuple[Any, ...]]) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 3 or tokens[0].upper() != "DELETE" or tokens[1].upper() != "FROM":
        raise ValueError(f"Invalid DELETE format: {query}")
    table = tokens[2]

    where = None
    if len(tokens) > 3:
        if tokens[3].upper() != "WHERE":
            raise ValueError(f"Invalid DELETE format: {query}")
        where_part = " ".join(tokens[4:])
        where = _parse_where_expression(
            where_part,
            params,
            allow_subqueries=True,
            outer_sources={table},
        )

    return {
        "action": "DELETE",
        "table": table,
        "where": where,
    }
