from typing import Any, Dict, List

from pandas import DataFrame


def execute_query(
    parsed: Dict[str, Any], data: Dict[str, DataFrame]
) -> List[Dict[str, Any]]:
    """
    Execute the parsed SQL query against the provided Excel data.

    Args:
        parsed (Dict[str, Any]): Parsed SQL query components.
        data (Dict[str, DataFrame]): Excel sheet data as DataFrames.

    Returns:
        List[Dict[str, Any]]: Query result as list of dictionaries.

    Raises:
        ValueError: If the specified sheet (table) is not found.
        NotImplementedError: For unsupported SQL actions.
    """
    table = parsed.get("table")
    if table not in data:
        raise ValueError(f"Sheet '{table}' not found in Excel")

    if parsed["action"] == "SELECT":
        df = data[table]
        if parsed.get("where"):
            try:
                df = df.query(parsed["where"])
            except Exception as e:
                raise ValueError(f"Invalid WHERE condition: {parsed['where']}") from e
        return df.to_dict(orient="records")

    elif parsed["action"] == "INSERT":
        raise NotImplementedError("INSERT is not yet implemented")

    else:
        raise NotImplementedError(f"Unsupported action: {parsed['action']}")
