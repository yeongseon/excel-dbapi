import re
from typing import Any, Dict, List, Optional


def _split_csv(text: str) -> List[str]:
    items: List[str] = []
    current: List[str] = []
    in_single = False
    in_double = False
    for char in text:
        if char == "'" and not in_double:
            in_single = not in_single
        elif char == '"' and not in_single:
            in_double = not in_double
        if char == "," and not in_single and not in_double:
            items.append("".join(current).strip())
            current = []
            continue
        current.append(char)
    if current:
        items.append("".join(current).strip())
    return items


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


def _parse_value(token: str) -> Any:
    token = token.strip()
    if token.upper() == "NULL":
        return None
    if token.startswith("'") and token.endswith("'") and len(token) >= 2:
        # Unescape doubled single quotes: 'it''s' -> it's
        return token[1:-1].replace("''", "'")
    if token.startswith('"') and token.endswith('"') and len(token) >= 2:
        # Unescape doubled double quotes: "say ""hello""" -> say "hello"
        return token[1:-1].replace('""', '"')
    try:
        return int(token)
    except ValueError:
        pass
    try:
        return float(token)
    except ValueError:
        return token


_AGGREGATE_FUNCTIONS = frozenset({"COUNT", "SUM", "AVG", "MIN", "MAX"})


def _collapse_aggregate_tokens(tokens: List[str]) -> List[str]:
    collapsed: List[str] = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        upper = token.upper()
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
    return (
        len(token) >= 2
        and ((token.startswith("'") and token.endswith("'")) or (token.startswith('"') and token.endswith('"')))
    )


def _find_clause_positions(tokens: List[str]) -> Dict[str, int]:
    positions: Dict[str, int] = {}
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if _is_quoted_token(token):
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


def _parse_columns(columns_token: str) -> List[Any]:
    columns_token = columns_token.strip()
    if columns_token == "*":
        return ["*"]
    columns: List[Any] = []
    for raw_column in _split_csv(columns_token):
        column = raw_column.strip()
        if not column:
            continue
        if re.search(r"(?i)\bOVER\s*\(", column):
            raise ValueError("Unsupported SQL syntax: OVER")
        match = re.fullmatch(
            r"(?i)(COUNT|SUM|AVG|MIN|MAX)\s*\(\s*([^\)]+?)\s*\)", column
        )
        if match:
            func = match.group(1).upper()
            arg = match.group(2).strip()
            if not arg:
                raise ValueError("Invalid aggregate expression")
            if arg == "*" and func != "COUNT":
                raise ValueError(f"{func} does not support *")
            if arg != "*" and not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", arg):
                raise ValueError(
                    f"Unsupported aggregate expression: {func}({arg}). "
                    "Only bare column names and * are supported"
                )
            columns.append({"type": "aggregate", "func": func, "arg": arg})
            continue
        columns.append(column)
    if not columns:
        raise ValueError("Invalid column list")
    return columns


def _values_to_bind_from_condition(condition: Dict[str, Any]) -> List[Any]:
    operator = str(condition["operator"]).upper()
    if operator in {"IS", "IS NOT"}:
        return []
    if operator == "IN":
        return list(condition["value"])
    if operator == "BETWEEN":
        low, high = condition["value"]
        return [low, high]
    return [condition["value"]]


def _apply_bound_values_to_condition(
    condition: Dict[str, Any], bound_values: List[Any], offset: int
) -> int:
    operator = str(condition["operator"]).upper()
    if operator in {"IS", "IS NOT"}:
        return 0
    if operator == "IN":
        size = len(condition["value"])
        condition["value"] = tuple(bound_values[offset : offset + size])
        return size
    if operator == "BETWEEN":
        condition["value"] = (bound_values[offset], bound_values[offset + 1])
        return 2
    condition["value"] = bound_values[offset]
    return 1


def _where_values_to_bind(where: Dict[str, Any]) -> List[Any]:
    values: List[Any] = []
    for condition in where["conditions"]:
        values.extend(_values_to_bind_from_condition(condition))
    return values


def _bind_where_conditions(
    where: Dict[str, Any], bound_values: List[Any], offset: int
) -> int:
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
) -> Dict[str, Any]:
    tokens = _collapse_aggregate_tokens(_tokenize(where_part.strip()))
    if not allow_aggregates and any(
        re.fullmatch(r"(?i)(COUNT|SUM|AVG|MIN|MAX)\([^\)]+\)", token)
        for token in tokens
    ):
        raise ValueError(
            "Aggregate functions are not allowed in WHERE clause; use HAVING instead"
        )
    for token_index, token in enumerate(tokens):
        if token.startswith("("):
            if token_index == 0 or tokens[token_index - 1].upper() != "IN":
                raise ValueError(
                    "Unsupported SQL grammar: parenthesized expressions in WHERE clause"
                )
    if len(tokens) < 3:
        raise ValueError("Invalid WHERE clause format")
    conditions: List[Dict[str, Any]] = []
    conjunctions: List[str] = []
    index = 0
    while index < len(tokens):
        if index + 1 >= len(tokens):
            raise ValueError("Invalid WHERE clause format")
        column = tokens[index]
        operator = tokens[index + 1].upper()

        # Handle IS NULL / IS NOT NULL
        if operator == "IS":
            if index + 2 < len(tokens) and tokens[index + 2].upper() == "NOT":
                if index + 3 < len(tokens) and tokens[index + 3].upper() == "NULL":
                    conditions.append(
                        {"column": column, "operator": "IS NOT", "value": None}
                    )
                    index += 4
                else:
                    raise ValueError(
                        "Invalid WHERE clause format: expected NULL after IS NOT"
                    )
            elif index + 2 < len(tokens) and tokens[index + 2].upper() == "NULL":
                conditions.append({"column": column, "operator": "IS", "value": None})
                index += 3
            else:
                raise ValueError(
                    "Invalid WHERE clause format: expected NULL or NOT after IS"
                )
        elif operator == "BETWEEN":
            if index + 4 >= len(tokens):
                raise ValueError("Invalid WHERE clause format")
            if tokens[index + 3].upper() != "AND":
                raise ValueError(
                    "Invalid WHERE clause format: expected AND in BETWEEN clause"
                )
            low_value = _parse_value(tokens[index + 2])
            high_value = _parse_value(tokens[index + 4])
            conditions.append(
                {
                    "column": column,
                    "operator": "BETWEEN",
                    "value": (low_value, high_value),
                }
            )
            index += 5
        elif operator == "IN":
            if index + 2 >= len(tokens):
                raise ValueError("Invalid WHERE clause format")
            in_start = index + 2
            in_end = in_start
            while in_end < len(tokens) and ")" not in tokens[in_end]:
                in_end += 1
            if in_end >= len(tokens):
                raise ValueError(
                    "Invalid WHERE clause format: expected ')' in IN clause"
                )

            in_values_text = " ".join(tokens[in_start : in_end + 1]).strip()
            if not in_values_text.startswith("(") or not in_values_text.endswith(")"):
                raise ValueError("Invalid WHERE clause format: malformed IN clause")

            raw_values = in_values_text[1:-1].strip()
            if not raw_values:
                raise ValueError(
                    "Invalid WHERE clause format: IN clause cannot be empty"
                )

            parsed_values = tuple(
                _parse_value(token) for token in _split_csv(raw_values)
            )
            if len(parsed_values) == 0:
                raise ValueError(
                    "Invalid WHERE clause format: IN clause cannot be empty"
                )

            conditions.append(
                {
                    "column": column,
                    "operator": "IN",
                    "value": parsed_values,
                }
            )
            index = in_end + 1
        elif operator == "LIKE":
            if index + 2 >= len(tokens):
                raise ValueError("Invalid WHERE clause format")
            value = _parse_value(tokens[index + 2])
            conditions.append({"column": column, "operator": "LIKE", "value": value})
            index += 3
        else:
            if index + 2 >= len(tokens):
                raise ValueError("Invalid WHERE clause format")
            value = _parse_value(tokens[index + 2])
            conditions.append(
                {"column": column, "operator": tokens[index + 1], "value": value}
            )
            index += 3

        if index < len(tokens):
            conj = tokens[index].upper()
            if conj not in {"AND", "OR"}:
                raise ValueError("Invalid WHERE clause format")
            conjunctions.append(conj)
            index += 1

    where_expression = {"conditions": conditions, "conjunctions": conjunctions}
    values_to_bind = _where_values_to_bind(where_expression)
    if bind_params and (
        params is not None or any(value == "?" for value in values_to_bind)
    ):
        bound = _bind_params(values_to_bind, params)
        _bind_where_conditions(where_expression, bound, 0)

    return where_expression


def _parse_select(query: str, params: Optional[tuple[Any, ...]]) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    from_index = -1
    for i, token in enumerate(tokens):
        if token.upper() == "FROM":
            from_index = i
            break
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
    columns = _parse_columns(columns_token)

    if len(tokens) <= from_index + 1:
        raise ValueError(f"Invalid SQL query format: {query}")
    table = tokens[from_index + 1]

    clause_tokens = tokens[from_index + 2 :]
    for token in clause_tokens:
        if _is_quoted_token(token):
            continue
        if token.upper() == "JOIN":
            raise ValueError("Unsupported SQL grammar:JOIN")
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
            for idx in [group_index, having_index, order_index, limit_index, offset_index]
            if idx >= 0 and idx > where_index
        ]
        where_end = min(where_end_candidates) if where_end_candidates else len(clause_tokens)
        where_part = " ".join(clause_tokens[where_start:where_end]).strip()
        where = _parse_where_expression(where_part, params, bind_params=False)

    if group_index >= 0:
        group_start = group_index + 2
        group_end_candidates = [
            idx for idx in [having_index, order_index, limit_index, offset_index] if idx >= 0 and idx > group_index
        ]
        group_end = min(group_end_candidates) if group_end_candidates else len(clause_tokens)
        group_part = " ".join(clause_tokens[group_start:group_end]).strip()
        group_columns = [col.strip() for col in _split_csv(group_part) if col.strip()]
        if not group_columns:
            raise ValueError("Invalid GROUP BY clause format")
        group_by = group_columns

    if having_index >= 0:
        having_start = having_index + 1
        having_end_candidates = [
            idx for idx in [order_index, limit_index, offset_index] if idx >= 0 and idx > having_index
        ]
        having_end = min(having_end_candidates) if having_end_candidates else len(clause_tokens)
        having_part = " ".join(clause_tokens[having_start:having_end]).strip()
        if not having_part:
            raise ValueError("Invalid HAVING clause format")
        having_part = _normalize_aggregate_expressions(having_part)
        having = _parse_where_expression(
            having_part,
            params,
            bind_params=False,
            allow_aggregates=True,
        )

    if order_index >= 0:
        order_start = order_index + 2
        order_end_candidates = [
            idx
            for idx in [limit_index, offset_index]
            if idx >= 0 and idx > order_index
        ]
        order_end = min(order_end_candidates) if order_end_candidates else len(clause_tokens)
        order_part = " ".join(clause_tokens[order_start:order_end]).strip()
        order_part = _normalize_aggregate_expressions(order_part)
        order_tokens = order_part.split()
        if not order_tokens:
            raise ValueError("Invalid ORDER BY clause format")
        direction = "ASC"
        if len(order_tokens) > 1:
            direction = order_tokens[1].upper()
        if direction not in {"ASC", "DESC"}:
            raise ValueError("Invalid ORDER BY direction")
        if len(order_tokens) > 2:
            raise ValueError(
                f"Unsupported SQL syntax: {' '.join(order_tokens[2:])}"
            )
        order_by = {"column": order_tokens[0], "direction": direction}

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
            for idx in [group_index, having_index, order_index, limit_index, offset_index]
            if idx >= 0 and idx > where_index
        ]
        where_end = min(where_end_candidates) if where_end_candidates else len(clause_tokens)
        consumed_indices.update(range(where_start, where_end))

    if group_index >= 0:
        group_start = group_index + 2
        group_end_candidates = [
            idx
            for idx in [having_index, order_index, limit_index, offset_index]
            if idx >= 0 and idx > group_index
        ]
        group_end = min(group_end_candidates) if group_end_candidates else len(clause_tokens)
        consumed_indices.update(range(group_start, group_end))

    if having_index >= 0:
        having_start = having_index + 1
        having_end_candidates = [
            idx for idx in [order_index, limit_index, offset_index] if idx >= 0 and idx > having_index
        ]
        having_end = min(having_end_candidates) if having_end_candidates else len(clause_tokens)
        consumed_indices.update(range(having_start, having_end))

    if order_index >= 0:
        order_start = order_index + 2
        order_end_candidates = [
            idx for idx in [limit_index, offset_index] if idx >= 0 and idx > order_index
        ]
        order_end = min(order_end_candidates) if order_end_candidates else len(clause_tokens)
        consumed_indices.update(range(order_start, order_end))

    if limit_index >= 0:
        limit_start = limit_index + 1
        limit_end = offset_index if offset_index >= 0 and offset_index > limit_index else len(clause_tokens)
        consumed_indices.update(range(limit_start, limit_end))

    if offset_index >= 0:
        consumed_indices.update(range(offset_index + 1, len(clause_tokens)))

    unconsumed = [i for i in range(len(clause_tokens)) if i not in consumed_indices]
    if unconsumed:
        unconsumed_text = " ".join(clause_tokens[i] for i in unconsumed)
        raise ValueError(f"Unsupported SQL syntax: {unconsumed_text}")

    if params is not None or (
        (where and any(value == "?" for value in _where_values_to_bind(where)))
        or (having and any(value == "?" for value in _where_values_to_bind(having)))
        or limit == "?"
        or offset == "?"
    ):
        values_to_bind = []
        if where:
            values_to_bind.extend(_where_values_to_bind(where))
        if having:
            values_to_bind.extend(_where_values_to_bind(having))
        if limit is not None:
            values_to_bind.append(limit)
        if offset is not None:
            values_to_bind.append(offset)
        bound = _bind_params(values_to_bind, params)
        consumed = 0
        if where:
            consumed += _bind_where_conditions(where, bound, consumed)
        if having:
            consumed += _bind_where_conditions(having, bound, consumed)
        if limit is not None:
            limit = bound[consumed]
            consumed += 1
        if offset is not None:
            offset = bound[consumed]

    if limit is not None and not isinstance(limit, int):
        raise ValueError("LIMIT must be an integer")
    if offset is not None and not isinstance(offset, int):
        raise ValueError("OFFSET must be an integer")

    return {
        "action": "SELECT",
        "columns": columns,
        "table": table,
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
        if any(value == "?" for value in values):
            raise ValueError("Missing parameters for placeholders")
        return values
    bound: List[Any] = []
    param_index = 0
    for value in values:
        if value == "?":
            if param_index >= len(params):
                raise ValueError("Not enough parameters for placeholders")
            bound.append(params[param_index])
            param_index += 1
        else:
            bound.append(value)
    if param_index < len(params):
        raise ValueError("Too many parameters for placeholders")
    return bound


def _parse_insert(query: str, params: Optional[tuple[Any, ...]]) -> Dict[str, Any]:
    upper = query.upper()
    if " VALUES " not in upper:
        raise ValueError(f"Invalid INSERT format: {query}")
    values_index = upper.index(" VALUES ")
    before_values = query[:values_index]
    values_part = query[values_index + len(" VALUES ") :]
    before_tokens = before_values.strip().split()
    if (
        len(before_tokens) < 3
        or before_tokens[0].upper() != "INSERT"
        or before_tokens[1].upper() != "INTO"
    ):
        raise ValueError(f"Invalid INSERT format: {query}")
    prefix_len = len(before_tokens[0]) + 1 + len(before_tokens[1])
    table_and_cols = before_values.strip()[prefix_len:].strip()

    columns = None
    if "(" in table_and_cols:
        table_name, cols_part = table_and_cols.split("(", 1)
        table = table_name.strip()
        cols_part = cols_part.rsplit(")", 1)[0]
        columns = _parse_columns(cols_part)
    else:
        table = table_and_cols.strip()

    values_part = values_part.strip()
    if not values_part.startswith("(") or not values_part.endswith(")"):
        raise ValueError(f"Invalid INSERT format: {query}")
    raw_values = values_part[1:-1].strip()
    values = [_parse_value(token) for token in _split_csv(raw_values)]
    values = _bind_params(values, params)

    return {
        "action": "INSERT",
        "table": table,
        "columns": columns,
        "values": values,
    }


def _parse_create(query: str) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 3 or tokens[0].upper() != "CREATE" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    table_and_cols = " ".join(tokens[2:]).strip()
    if "(" not in table_and_cols or not table_and_cols.endswith(")"):
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    table_name, cols_part = table_and_cols.split("(", 1)
    table = table_name.strip()
    cols_part = cols_part.rsplit(")", 1)[0]
    raw_columns = _split_csv(cols_part)
    columns = []
    for col in raw_columns:
        if not col:
            continue
        columns.append(col.strip().split()[0])
    if not columns:
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    return {
        "action": "CREATE",
        "table": table,
        "columns": columns,
    }


def _parse_drop(query: str) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) != 3 or tokens[0].upper() != "DROP" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid DROP TABLE format: {query}")
    return {
        "action": "DROP",
        "table": tokens[2],
    }


def parse_sql(query: str, params: Optional[tuple[Any, ...]] = None) -> Dict[str, Any]:
    tokens = query.strip().split()
    if not tokens:
        raise ValueError(f"Invalid SQL query format: {query}")
    action = tokens[0].upper()
    if action == "SELECT":
        parsed = _parse_select(query, params)
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
        assignments.append({"column": col.strip(), "value": _parse_value(value)})

    where = None
    if where_part:
        where = _parse_where_expression(where_part, params, bind_params=False)

    values_to_bind = [item["value"] for item in assignments]
    if where is not None:
        values_to_bind.extend(_where_values_to_bind(where))
    if params is not None or any(value == "?" for value in values_to_bind):
        bound = _bind_params(values_to_bind, params)
        for idx, item in enumerate(assignments):
            item["value"] = bound[idx]
        if where is not None:
            offset = len(assignments)
            _bind_where_conditions(where, bound, offset)

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
        where = _parse_where_expression(where_part, params)

    return {
        "action": "DELETE",
        "table": table,
        "where": where,
    }
