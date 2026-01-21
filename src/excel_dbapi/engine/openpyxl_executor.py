from typing import Any, Dict, List, Sequence

from .result import ExecutionResult, Description

class OpenpyxlExecutor:
    """
    OpenpyxlExecutor is responsible for executing parsed SQL-like queries
    on in-memory Excel worksheet data using openpyxl.
    """

    def __init__(self, data: Dict[str, Any]):
        """
        Initialize the OpenpyxlExecutor.

        Args:
            data (Dict[str, Any]): A dictionary mapping sheet names to openpyxl Worksheet objects.
        """
        self.data = data

    def execute(self, parsed: Dict[str, Any]) -> ExecutionResult:
        """
        Execute a parsed SQL-like query on the in-memory Excel data.

        Args:
            parsed (Dict[str, Any]): A parsed query dictionary containing keys like
                                     'table', 'columns', 'where', etc.

        Returns:
            List[Dict[str, Any]]: Query results as a list of dictionaries, where each dictionary
                                  represents a row with column names as keys.

        Raises:
            ValueError: If the specified table (sheet) does not exist.
            NotImplementedError: If an unsupported operator is used in the WHERE clause.
        """
        table = parsed["table"]
        ws = self.data.get(table)
        if ws is None:
            raise ValueError(f"Sheet '{table}' not found in Excel")

        # Read all rows from the worksheet
        rows = list(ws.iter_rows(values_only=True))
        headers = list(rows[0])  # First row is assumed to be the header
        data = [dict(zip(headers, row)) for row in rows[1:]]

        columns: Sequence[str] = parsed["columns"]
        if columns == ["*"]:
            selected_columns = headers
        else:
            selected_columns = list(columns)
            missing = [col for col in selected_columns if col not in headers]
            if missing:
                raise ValueError(f"Unknown column(s): {', '.join(missing)}")

        # Apply WHERE clause filtering if provided
        where = parsed.get("where")
        if where:
            column = where["column"]
            operator = where["operator"]
            value = where["value"]

            if operator == "=":
                data = [row for row in data if str(row.get(column)) == str(value)]
            else:
                raise NotImplementedError(f"Unsupported operator: {operator}")

        rows_out = [tuple(row.get(col) for col in selected_columns) for row in data]
        description: Description = [
            (col, None, None, None, None, None, None) for col in selected_columns
        ]

        return ExecutionResult(
            rows=rows_out,
            description=description,
            rowcount=len(rows_out),
            lastrowid=None,
        )
