import re
from typing import Any, Dict, Optional, Tuple


def parse_sql(query: str, params: Optional[Tuple[Any, ...]] = None) -> Dict[str, Any]:
    query = query.strip()

    # Parameter binding
    if params:
        placeholders = re.findall(r"\?", query)
        if len(placeholders) != len(params):
            raise ValueError(
                f"Expected {len(placeholders)} parameters, got {len(params)}"
            )

        for param in params:
            if isinstance(param, str):
                value = f"'{param}'"
            else:
                value = str(param)
            query = query.replace("?", value, 1)

    result: Dict[str, Any] = {}
    lower_query = query.lower()

    if lower_query.startswith("select"):
        result["action"] = "SELECT"
        table_match = re.search(r"from\s+\[?(\w+)\$?\]?", query, re.IGNORECASE)
        result["table"] = table_match.group(1).lower() if table_match else None

        where_match = re.search(r"where\s+(.+)", query, re.IGNORECASE)
        if where_match:
            condition = where_match.group(1).strip()
            # SQL → pandas 변환
            condition = re.sub(r"(?<![=!<>])=(?!=)", "==", condition)
            result["where"] = condition
        else:
            result["where"] = None

    elif lower_query.startswith("insert"):
        result["action"] = "INSERT"
        table_match = re.search(r"into\s+\[?(\w+)\$?\]?", query, re.IGNORECASE)
        result["table"] = table_match.group(1).lower() if table_match else None

    else:
        raise NotImplementedError(f"Unsupported SQL: {query}")

    result["query"] = query
    return result
