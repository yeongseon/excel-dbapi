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


def _parse_value(token: str) -> Any:
    token = token.strip()
    if token.upper() == "NULL":
        return None
    if token.startswith(("'", '"')) and token.endswith(("'", '"')):
        return token[1:-1]
    try:
        return int(token)
    except ValueError:
        pass
    try:
        return float(token)
    except ValueError:
        return token


def _parse_columns(columns_token: str) -> List[str]:
    columns_token = columns_token.strip()
    if columns_token == "*":
        return ["*"]
    columns = [col.strip() for col in columns_token.split(",") if col.strip()]
    if not columns:
        raise ValueError("Invalid column list")
    return columns


def _parse_select(query: str, params: Optional[tuple]) -> Dict[str, Any]:
    tokens = query.strip().split()
    try:
        from_index = tokens.index("FROM")
    except ValueError:
        try:
            from_index = tokens.index("from")
        except ValueError as exc:
            raise ValueError(f"Invalid SQL query format: {query}") from exc

    columns_token = " ".join(tokens[1:from_index]).strip()
    if not columns_token:
        raise ValueError(f"Invalid SQL query format: {query}")
    columns = _parse_columns(columns_token)

    if len(tokens) <= from_index + 1:
        raise ValueError(f"Invalid SQL query format: {query}")
    table = tokens[from_index + 1]

    where = None
    if len(tokens) > from_index + 2 and tokens[from_index + 2].upper() == "WHERE":
        if len(tokens) < from_index + 6:
            raise ValueError(f"Invalid WHERE clause format: {query}")
        column = tokens[from_index + 3]
        operator = tokens[from_index + 4]
        value = _parse_value(tokens[from_index + 5])
        if params is not None or value == "?":
            value = _bind_params([value], params)[0]
        where = {
            "column": column,
            "operator": operator,
            "value": value,
        }

    return {
        "action": "SELECT",
        "columns": columns,
        "table": table,
        "where": where,
    }


def _bind_params(values: List[Any], params: Optional[tuple]) -> List[Any]:
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


def _parse_insert(query: str, params: Optional[tuple]) -> Dict[str, Any]:
    upper = query.upper()
    if " VALUES " not in upper:
        raise ValueError(f"Invalid INSERT format: {query}")
    values_index = upper.index(" VALUES ")
    before_values = query[:values_index]
    values_part = query[values_index + len(" VALUES "):]
    before_tokens = before_values.strip().split()
    if len(before_tokens) < 3 or before_tokens[0].upper() != "INSERT" or before_tokens[1].upper() != "INTO":
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
    tokens = query.strip().split(None, 2)
    if len(tokens) < 3 or tokens[0].upper() != "CREATE" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    table_and_cols = tokens[2].strip()
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
    tokens = query.strip().split()
    if len(tokens) != 3 or tokens[0].upper() != "DROP" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid DROP TABLE format: {query}")
    return {
        "action": "DROP",
        "table": tokens[2],
    }


def parse_sql(query: str, params: Optional[tuple] = None) -> Dict[str, Any]:
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


def _parse_update(query: str, params: Optional[tuple]) -> Dict[str, Any]:
    upper = query.upper()
    if " SET " not in upper:
        raise ValueError(f"Invalid UPDATE format: {query}")
    set_index = upper.index(" SET ")
    before_set = query[:set_index]
    after_set = query[set_index + len(" SET "):]
    before_tokens = before_set.strip().split()
    if len(before_tokens) < 2 or before_tokens[0].upper() != "UPDATE":
        raise ValueError(f"Invalid UPDATE format: {query}")
    table = before_tokens[1].strip()

    where_part = None
    after_upper = after_set.upper()
    if " WHERE " in after_upper:
        where_index = after_upper.index(" WHERE ")
        set_part = after_set[:where_index]
        where_part = after_set[where_index + len(" WHERE "):]
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
        where_tokens = where_part.strip().split()
        if len(where_tokens) < 3:
            raise ValueError(f"Invalid WHERE clause format: {query}")
        where = {
            "column": where_tokens[0],
            "operator": where_tokens[1],
            "value": _parse_value(where_tokens[2]),
        }

    values_to_bind = [item["value"] for item in assignments]
    if where is not None:
        values_to_bind.append(where["value"])
    if params is not None or any(value == "?" for value in values_to_bind):
        bound = _bind_params(values_to_bind, params)
        for idx, item in enumerate(assignments):
            item["value"] = bound[idx]
        if where is not None:
            where["value"] = bound[-1]

    return {
        "action": "UPDATE",
        "table": table,
        "set": assignments,
        "where": where,
    }


def _parse_delete(query: str, params: Optional[tuple]) -> Dict[str, Any]:
    tokens = query.strip().split()
    if len(tokens) < 3 or tokens[0].upper() != "DELETE" or tokens[1].upper() != "FROM":
        raise ValueError(f"Invalid DELETE format: {query}")
    table = tokens[2]

    where = None
    if len(tokens) > 3:
        if tokens[3].upper() != "WHERE" or len(tokens) < 7:
            raise ValueError(f"Invalid DELETE format: {query}")
        where = {
            "column": tokens[4],
            "operator": tokens[5],
            "value": _parse_value(tokens[6]),
        }

    if where is not None and (params is not None or where["value"] == "?"):
        where["value"] = _bind_params([where["value"]], params)[0]

    return {
        "action": "DELETE",
        "table": table,
        "where": where,
    }
