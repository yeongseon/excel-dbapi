from typing import Any, Dict, List, Optional


def parse_sql(query: str, params: Optional[tuple] = None) -> Dict[str, Any]:
    """
    Parse a simple SQL SELECT query with optional WHERE clause.
    Supports queries like:
      SELECT * FROM Sheet1
      SELECT * FROM Sheet1 WHERE id = 1
    """
    tokens = query.strip().split()

    if len(tokens) < 4 or tokens[0].upper() != "SELECT":
        raise ValueError(f"Invalid SQL query format: {query}")

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

    if columns_token == "*":
        columns: List[str] = ["*"]
    else:
        columns = [col.strip() for col in columns_token.split(",") if col.strip()]
        if not columns:
            raise ValueError(f"Invalid SQL query format: {query}")

    table = tokens[from_index + 1]

    # Check if there's a WHERE clause
    where = None
    if len(tokens) > from_index + 2 and tokens[from_index + 2].upper() == "WHERE":
        # For now, assume very simple: WHERE column = value
        if len(tokens) < from_index + 6:
            raise ValueError(f"Invalid WHERE clause format: {query}")
        column = tokens[from_index + 3]
        operator = tokens[from_index + 4]
        value = tokens[from_index + 5]

        # Remove quotes if present
        if value.startswith(("'", '"')) and value.endswith(("'", '"')):
            value = value[1:-1]

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
        "params": params,
    }
