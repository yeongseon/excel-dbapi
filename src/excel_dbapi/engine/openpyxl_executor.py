from typing import Any, Dict, List, Sequence

from .result import ExecutionResult, Description

class OpenpyxlExecutor:
    """
    OpenpyxlExecutor is responsible for executing parsed SQL-like queries
    on in-memory Excel worksheet data using openpyxl.
    """

    def __init__(self, data: Dict[str, Any], workbook: Any):
        """
        Initialize the OpenpyxlExecutor.

        Args:
            data (Dict[str, Any]): A dictionary mapping sheet names to openpyxl Worksheet objects.
        """
        self.data = data
        self.workbook = workbook

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
        action = parsed["action"]
        table = parsed["table"]

        if action == "SELECT":
            ws = self.data.get(table)
            if ws is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return ExecutionResult(action=action, rows=[], description=[], rowcount=0, lastrowid=None)
            headers = list(rows[0])
            data = [dict(zip(headers, row)) for row in rows[1:]]

            columns = parsed["columns"]
            if columns == ["*"]:
                selected_columns = headers
            else:
                selected_columns = list(columns)
                missing = [col for col in selected_columns if col not in headers]
                if missing:
                    raise ValueError(f"Unknown column(s): {', '.join(missing)}")

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
                action=action,
                rows=rows_out,
                description=description,
                rowcount=len(rows_out),
                lastrowid=None,
            )

        if action == "UPDATE":
            ws = self.data.get(table)
            if ws is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return ExecutionResult(action=action, rows=[], description=[], rowcount=0, lastrowid=None)
            headers = list(rows[0])
            updates = parsed["set"]
            for update in updates:
                if update["column"] not in headers:
                    raise ValueError(f"Unknown column: {update['column']}")

            where = parsed.get("where")
            rowcount = 0
            for row_index in range(2, ws.max_row + 1):
                row_values = {
                    headers[col_index]: ws.cell(row=row_index, column=col_index + 1).value
                    for col_index in range(len(headers))
                }
                if where and not self._matches_where(row_values, where):
                    continue
                for update in updates:
                    col_index = headers.index(update["column"]) + 1
                    ws.cell(row=row_index, column=col_index, value=update["value"])
                rowcount += 1

            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=rowcount,
                lastrowid=None,
            )

        if action == "DELETE":
            ws = self.data.get(table)
            if ws is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return ExecutionResult(action=action, rows=[], description=[], rowcount=0, lastrowid=None)
            headers = list(rows[0])
            where = parsed.get("where")
            rowcount = 0
            for row_index in range(ws.max_row, 1, -1):
                row_values = {
                    headers[col_index]: ws.cell(row=row_index, column=col_index + 1).value
                    for col_index in range(len(headers))
                }
                if where and not self._matches_where(row_values, where):
                    continue
                ws.delete_rows(row_index)
                rowcount += 1

            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=rowcount,
                lastrowid=None,
            )

        if action == "INSERT":
            ws = self.data.get(table)
            if ws is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                raise ValueError("Cannot insert into sheet without headers")
            headers = list(rows[0])

            values = parsed["values"]
            insert_columns = parsed.get("columns")
            if insert_columns is None:
                if len(values) != len(headers):
                    raise ValueError("INSERT values count does not match header count")
                row_values = list(values)
            else:
                missing = [col for col in insert_columns if col not in headers]
                if missing:
                    raise ValueError(f"Unknown column(s): {', '.join(missing)}")
                if len(values) != len(insert_columns):
                    raise ValueError("INSERT values count does not match column count")
                row_values = [None for _ in headers]
                for col, value in zip(insert_columns, values):
                    row_values[headers.index(col)] = value

            ws.append(row_values)
            last_row = ws.max_row
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=1,
                lastrowid=last_row,
            )

        if action == "CREATE":
            if self.workbook is None:
                raise ValueError("Workbook is not loaded")
            if table in self.data:
                raise ValueError(f"Sheet '{table}' already exists")
            ws = self.workbook.create_sheet(title=table)
            ws.append(parsed["columns"])
            self.data[table] = ws
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        if action == "DROP":
            if self.workbook is None:
                raise ValueError("Workbook is not loaded")
            ws = self.data.get(table)
            if ws is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            self.workbook.remove(ws)
            del self.data[table]
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        raise ValueError(f"Unsupported action: {action}")

    def _matches_where(self, row: Dict[str, Any], where: Dict[str, Any]) -> bool:
        column = where["column"]
        operator = where["operator"]
        value = where["value"]
        if operator == "=":
            return str(row.get(column)) == str(value)
        raise NotImplementedError(f"Unsupported operator: {operator}")
