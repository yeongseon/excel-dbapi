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
                data = [row for row in data if self._matches_where(row, where)]

            order_by = parsed.get("order_by")
            if order_by:
                if order_by["column"] not in headers:
                    raise ValueError(f"Unknown column: {order_by['column']}")
                reverse = order_by["direction"] == "DESC"
                data = sorted(
                    data,
                    key=lambda row: self._sort_key(row.get(order_by["column"])),
                    reverse=reverse,
                )

            limit = parsed.get("limit")
            if limit is not None:
                data = data[:limit]

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
        if "conditions" in where:
            conditions = where["conditions"]
            conjunctions = where["conjunctions"]
            result = self._evaluate_condition(row, conditions[0])
            for idx, conj in enumerate(conjunctions):
                next_result = self._evaluate_condition(row, conditions[idx + 1])
                if conj == "AND":
                    result = result and next_result
                else:
                    result = result or next_result
            return result

        return self._evaluate_condition(row, where)

    def _evaluate_condition(self, row: Dict[str, Any], condition: Dict[str, Any]) -> bool:
        column = condition["column"]
        operator = condition["operator"]
        value = condition["value"]
        row_value = row.get(column)

        left, right = self._coerce_for_compare(row_value, value)
        if operator in {"=", "=="}:
            return left == right
        if operator in {"!=", "<>"}:
            return left != right
        if operator == ">":
            return left > right
        if operator == ">=":
            return left >= right
        if operator == "<":
            return left < right
        if operator == "<=":
            return left <= right
        raise NotImplementedError(f"Unsupported operator: {operator}")

    def _coerce_for_compare(self, left: Any, right: Any) -> tuple[Any, Any]:
        left_num = self._to_number(left)
        right_num = self._to_number(right)
        if left_num is not None and right_num is not None:
            return left_num, right_num
        return str(left), str(right)

    def _sort_key(self, value: Any) -> tuple[int, Any]:
        if value is None:
            return (1, "")
        numeric = self._to_number(value)
        if numeric is not None:
            return (0, numeric)
        return (0, str(value))

    def _to_number(self, value: Any) -> float | None:
        if isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            try:
                return float(value)
            except ValueError:
                return None
        return None
