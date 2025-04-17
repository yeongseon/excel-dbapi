from typing import Any, Dict, List

class OpenpyxlExecutor:
    def __init__(self, data: Dict[str, Any]):
        self.data = data

    def execute(self, parsed: Dict[str, Any]) -> List[Dict[str, Any]]:
        table = parsed["table"]
        ws = self.data.get(table)
        if ws is None:
            raise ValueError(f"Sheet '{table}' not found in Excel")

        rows = list(ws.iter_rows(values_only=True))
        headers = rows[0]
        return [dict(zip(headers, row)) for row in rows[1:]]
