from typing import Any, Dict, Optional


def parse_sql(query: str, params: Optional[tuple] = None) -> Dict[str, Any]:
    """
    Parse a simple SQL SELECT query with optional WHERE clause.
    Supports queries like:
      SELECT * FROM Sheet1
      SELECT * FROM Sheet1 WHERE id = 1
    """
    tokens = query.strip().split()

    if len(tokens) < 4 or tokens[0].upper() != "SELECT" or tokens[2].upper() != "FROM":
        raise ValueError(f"Invalid SQL query format: {query}")

    table = tokens[3]

    # Check if there's a WHERE clause
    where = None
    if len(tokens) > 4 and tokens[4].upper() == "WHERE":
        # For now, assume very simple: WHERE column = value
        if len(tokens) < 8:
            raise ValueError(f"Invalid WHERE clause format: {query}")
        column = tokens[5]
        operator = tokens[6]
        value = tokens[7]

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
        "columns": tokens[1],
        "table": table,
        "where": where,
        "params": params,
    }
