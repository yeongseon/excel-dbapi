from typing import Any, Dict, Sequence

import pandas as pd

from .result import Description, ExecutionResult


class PandasExecutor:
    def __init__(self, data: Dict[str, Any]):
        self.data = data

    def execute(self, parsed: Dict[str, Any]) -> ExecutionResult:
        action = parsed["action"]
        table = parsed["table"]

        if action == "SELECT":
            frame = self.data.get(table)
            if frame is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")

            columns: Sequence[str] = parsed["columns"]
            if columns == ["*"]:
                selected = frame
                selected_columns = list(frame.columns)
            else:
                missing = [col for col in columns if col not in frame.columns]
                if missing:
                    raise ValueError(f"Unknown column(s): {', '.join(missing)}")
                selected = frame[list(columns)]
                selected_columns = list(columns)

            where = parsed.get("where")
            if where:
                column = where["column"]
                operator = where["operator"]
                value = where["value"]
                if column not in selected.columns:
                    raise ValueError(f"Unknown column: {column}")
                if operator == "=":
                    mask = selected[column].astype(str) == str(value)
                else:
                    raise NotImplementedError(f"Unsupported operator: {operator}")
                selected = selected[mask]

            rows_out = list(selected.itertuples(index=False, name=None))
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

        if action == "INSERT":
            frame = self.data.get(table)
            if frame is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")

            values = parsed["values"]
            columns = parsed.get("columns")
            if columns is None:
                if len(values) != len(frame.columns):
                    raise ValueError("INSERT values count does not match header count")
                row_data = dict(zip(frame.columns, values))
            else:
                missing = [col for col in columns if col not in frame.columns]
                if missing:
                    raise ValueError(f"Unknown column(s): {', '.join(missing)}")
                row_data = {col: None for col in frame.columns}
                for col, value in zip(columns, values):
                    row_data[col] = value

            self.data[table] = pd.concat(
                [frame, pd.DataFrame([row_data])],
                ignore_index=True,
            )
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=1,
                lastrowid=len(self.data[table]),
            )

        if action == "CREATE":
            if table in self.data:
                raise ValueError(f"Sheet '{table}' already exists")
            self.data[table] = pd.DataFrame(columns=parsed["columns"])
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        if action == "DROP":
            if table not in self.data:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            del self.data[table]
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        raise ValueError(f"Unsupported action: {action}")
