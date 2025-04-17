from typing import Any, Dict, List

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

    def execute(self, parsed: Dict[str, Any]) -> List[Dict[str, Any]]:
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
        headers = rows[0]  # First row is assumed to be the header
        data = [dict(zip(headers, row)) for row in rows[1:]]

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

        return data
