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

            base = frame
            where = parsed.get("where")
            if where:
                mask = self._build_mask(base, where)
                base = base[mask]

            order_by = parsed.get("order_by")
            if order_by:
                if order_by["column"] not in base.columns:
                    raise ValueError(f"Unknown column: {order_by['column']}")
                base = base.sort_values(
                    by=order_by["column"],
                    ascending=order_by["direction"] == "ASC",
                )

            limit = parsed.get("limit")
            if limit is not None:
                base = base.head(limit)

            columns = parsed["columns"]
            if columns == ["*"]:
                selected = base
                selected_columns = list(base.columns)
            else:
                missing = [col for col in columns if col not in base.columns]
                if missing:
                    raise ValueError(f"Unknown column(s): {', '.join(missing)}")
                selected = base[list(columns)]
                selected_columns = list(columns)

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

        if action == "UPDATE":
            frame = self.data.get(table)
            if frame is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            where = parsed.get("where")
            if where:
                mask = self._build_mask(frame, where)
            else:
                mask = pd.Series([True] * len(frame), index=frame.index)

            updates = parsed["set"]
            for update in updates:
                if update["column"] not in frame.columns:
                    raise ValueError(f"Unknown column: {update['column']}")
                frame.loc[mask, update["column"]] = update["value"]
            self.data[table] = frame
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=int(mask.sum()),
                lastrowid=None,
            )

        if action == "DELETE":
            frame = self.data.get(table)
            if frame is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            where = parsed.get("where")
            if where:
                mask = self._build_mask(frame, where)
            else:
                mask = pd.Series([True] * len(frame), index=frame.index)

            rowcount = int(mask.sum())
            self.data[table] = frame.loc[~mask].reset_index(drop=True)
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=rowcount,
                lastrowid=None,
            )

        if action == "INSERT":
            frame = self.data.get(table)
            if frame is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")

            values = parsed["values"]
            insert_columns = parsed.get("columns")
            if insert_columns is None:
                if len(values) != len(frame.columns):
                    raise ValueError("INSERT values count does not match header count")
                row_data = dict(zip(frame.columns, values))
            else:
                missing = [col for col in insert_columns if col not in frame.columns]
                if missing:
                    raise ValueError(f"Unknown column(s): {', '.join(missing)}")
                if len(values) != len(insert_columns):
                    raise ValueError("INSERT values count does not match column count")
                row_data = {col: None for col in frame.columns}
                for col, value in zip(insert_columns, values):
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

    def _build_mask(self, frame: pd.DataFrame, where: Dict[str, Any]) -> pd.Series:
        if "conditions" in where:
            conditions = where["conditions"]
            conjunctions = where["conjunctions"]
            mask = self._evaluate_condition(frame, conditions[0])
            for idx, conj in enumerate(conjunctions):
                next_mask = self._evaluate_condition(frame, conditions[idx + 1])
                if conj == "AND":
                    mask = mask & next_mask
                else:
                    mask = mask | next_mask
            return mask
        return self._evaluate_condition(frame, where)

    def _evaluate_condition(self, frame: pd.DataFrame, condition: Dict[str, Any]) -> pd.Series:
        column = condition["column"]
        operator = condition["operator"]
        value = condition["value"]
        if column not in frame.columns:
            raise ValueError(f"Unknown column: {column}")
        series = frame[column]

        if operator in {"=", "=="}:
            return series == value
        if operator in {"!=", "<>"}:
            return series != value
        if operator == ">":
            return series > value
        if operator == ">=":
            return series >= value
        if operator == "<":
            return series < value
        if operator == "<=":
            return series <= value
        raise NotImplementedError(f"Unsupported operator: {operator}")
