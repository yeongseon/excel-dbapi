from typing import Any, Dict, List
from .openpyxl_executor import OpenpyxlExecutor


def execute_query(parsed: Dict[str, Any], data: Dict[str, Any]) -> List[Dict[str, Any]]:
    table = parsed.get("table").lower()
    data_lower = {sheet.lower(): sheet for sheet in data.keys()}

    if table not in data_lower:
        raise ValueError(f"Sheet '{table}' not found in Excel")

    actual_sheet_name = data_lower.get(table)
    return OpenpyxlExecutor(data).execute(parsed)
