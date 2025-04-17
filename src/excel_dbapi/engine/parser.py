from typing import Any, Dict, Optional


def parse_sql(query: str, params: Optional[tuple] = None) -> Dict[str, Any]:
    """
    Parse a very simple SQL SELECT query.
    Only supports 'SELECT * FROM table' style for now.
    """
    tokens = query.strip().split()
    if len(tokens) < 4 or tokens[0].upper() != "SELECT" or tokens[2].upper() != "FROM":
        raise ValueError(f"Invalid SQL query format: {query}")

    return {
        "action": "SELECT",
        "columns": tokens[1],  # usually '*'
        "table": tokens[3],
        "where": None,
        "params": params,
    }
