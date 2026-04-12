import re
from typing import Any

from .engines.base import WorkbookBackend
from .engines.result import Description, ExecutionResult
from .parser import parse_sql
from .sanitize import sanitize_cell_value, sanitize_row

_READONLY_ACTIONS = frozenset({"INSERT", "UPDATE", "DELETE", "CREATE", "DROP"})


class SharedExecutor:
    def __init__(self, backend: WorkbookBackend, *, sanitize_formulas: bool = True):
        self.backend = backend
        self.sanitize_formulas = sanitize_formulas

    def _ensure_writable(self, action: str) -> None:
        """Raise NotSupportedError if backend is read-only and action mutates data."""
        if action in _READONLY_ACTIONS and getattr(self.backend, "readonly", False):
            from .exceptions import NotSupportedError

            raise NotSupportedError(
                f"{action} is not supported by the read-only backend"
            )

    def execute_with_params(
        self, query: str, params: tuple[Any, ...] | None = None
    ) -> ExecutionResult:
        # Early readonly guard — extract SQL verb before full parse
        # so mutations are rejected before param-binding errors.
        first_word = query.strip().split(None, 1)[0].upper() if query.strip() else ""
        self._ensure_writable(first_word)
        parsed = parse_sql(query, params)
        return self.execute(parsed)

    def execute(self, parsed: dict[str, Any]) -> ExecutionResult:
        action = parsed["action"]
        self._ensure_writable(action)
        table = parsed["table"]
        resolved_table = self._resolve_sheet_name(table)

        if action == "SELECT":
            if resolved_table is None:
                available = self.backend.list_sheets()
                msg = f"Sheet '{table}' not found in Excel."
                if available:
                    msg += f" Available sheets: {available}"
                raise ValueError(msg)
            data = self.backend.read_sheet(resolved_table)
            if not data.headers:
                return ExecutionResult(
                    action=action, rows=[], description=[], rowcount=0, lastrowid=None
                )
            headers = list(data.headers)
            rows = [dict(zip(headers, row)) for row in data.rows]
            return self._execute_select(action, parsed, headers, rows)

        if action == "UPDATE":
            if resolved_table is None:
                available = self.backend.list_sheets()
                msg = f"Sheet '{table}' not found in Excel."
                if available:
                    msg += f" Available sheets: {available}"
                raise ValueError(msg)
            table_data = self.backend.read_sheet(resolved_table)
            if not table_data.headers:
                return ExecutionResult(
                    action=action, rows=[], description=[], rowcount=0, lastrowid=None
                )
            headers = list(table_data.headers)
            updates = parsed["set"]
            for update in updates:
                if update["column"] not in headers:
                    raise ValueError(
                        f"Unknown column: {update['column']}. Available columns: {headers}"
                    )

            where = parsed.get("where")
            rowcount = 0
            for row_values in table_data.rows:
                row_map = {
                    headers[col_index]: row_values[col_index]
                    if col_index < len(row_values)
                    else None
                    for col_index in range(len(headers))
                }
                if where and not self._matches_where(row_map, where):
                    continue
                for update in updates:
                    col_index = headers.index(update["column"])
                    value = (
                        sanitize_cell_value(update["value"])
                        if self.sanitize_formulas
                        else update["value"]
                    )
                    if col_index >= len(row_values):
                        row_values.extend([None] * (col_index - len(row_values) + 1))
                    row_values[col_index] = value
                rowcount += 1

            self.backend.write_sheet(resolved_table, table_data)
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=rowcount,
                lastrowid=None,
            )

        if action == "DELETE":
            if resolved_table is None:
                available = self.backend.list_sheets()
                msg = f"Sheet '{table}' not found in Excel."
                if available:
                    msg += f" Available sheets: {available}"
                raise ValueError(msg)
            table_data = self.backend.read_sheet(resolved_table)
            if not table_data.headers:
                return ExecutionResult(
                    action=action, rows=[], description=[], rowcount=0, lastrowid=None
                )
            headers = list(table_data.headers)
            where = parsed.get("where")
            rowcount = 0
            kept_rows: list[list[Any]] = []
            for row_values in table_data.rows:
                row_map = {
                    headers[col_index]: row_values[col_index]
                    if col_index < len(row_values)
                    else None
                    for col_index in range(len(headers))
                }
                if where and not self._matches_where(row_map, where):
                    kept_rows.append(row_values)
                    continue
                if where is None:
                    rowcount += 1
                else:
                    rowcount += 1
            table_data.rows = kept_rows
            self.backend.write_sheet(resolved_table, table_data)
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=rowcount,
                lastrowid=None,
            )

        if action == "INSERT":
            if resolved_table is None:
                available = self.backend.list_sheets()
                msg = f"Sheet '{table}' not found in Excel."
                if available:
                    msg += f" Available sheets: {available}"
                raise ValueError(msg)
            table_data = self.backend.read_sheet(resolved_table)
            if not table_data.headers:
                raise ValueError("Cannot insert into sheet without headers")
            headers = list(table_data.headers)

            values = parsed["values"]
            insert_columns = parsed.get("columns")
            if insert_columns is None:
                if len(values) != len(headers):
                    raise ValueError("INSERT values count does not match header count")
                row_values = list(values)
            else:
                missing = [col for col in insert_columns if col not in headers]
                if missing:
                    raise ValueError(
                        f"Unknown column(s): {', '.join(missing)}. Available columns: {headers}"
                    )
                if len(values) != len(insert_columns):
                    raise ValueError("INSERT values count does not match column count")
                row_values = [None for _ in headers]
                for col, value in zip(insert_columns, values):
                    row_values[headers.index(col)] = value

            sanitized_row = (
                sanitize_row(row_values) if self.sanitize_formulas else row_values
            )
            last_row = self.backend.append_row(resolved_table, sanitized_row)
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=1,
                lastrowid=last_row,
            )

        if action == "CREATE":
            if resolved_table is not None:
                raise ValueError(f"Sheet '{table}' already exists")
            self.backend.create_sheet(table, parsed["columns"])
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        if action == "DROP":
            if resolved_table is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            self.backend.drop_sheet(resolved_table)
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        raise ValueError(f"Unsupported action: {action}")

    def _matches_where(self, row: dict[str, Any], where: dict[str, Any]) -> bool:
        if "conditions" in where:
            conditions = where["conditions"]
            conjunctions = where["conjunctions"]
            results = [self._evaluate_condition(row, conditions[0])]
            for idx, conj in enumerate(conjunctions):
                next_result = self._evaluate_condition(row, conditions[idx + 1])
                if conj == "AND":
                    results[-1] = results[-1] and next_result
                else:
                    results.append(next_result)
            return any(results)

        return self._evaluate_condition(row, where)

    def _execute_select(
        self,
        action: str,
        parsed: dict[str, Any],
        headers: list[str],
        rows: list[dict[str, Any]],
    ) -> ExecutionResult:
        columns = parsed["columns"]
        where = parsed.get("where")
        if where:
            rows = [row for row in rows if self._matches_where(row, where)]

        group_by: list[str] | None = parsed.get("group_by")
        having = parsed.get("having")
        aggregate_query = any(self._is_aggregate_column(col) for col in columns)

        if columns == ["*"] and (aggregate_query or group_by):
            raise ValueError("SELECT * is not supported with GROUP BY or aggregate functions")

        if aggregate_query or group_by is not None:
            return self._execute_aggregate_select(
                action, parsed, headers, rows, columns, group_by, having
            )

        # --- Non-aggregate path ---
        selected_columns = headers if columns == ["*"] else list(columns)
        missing = [col for col in selected_columns if col not in headers]
        if missing:
            raise ValueError(
                f"Unknown column(s): {', '.join(missing)}. Available columns: {headers}"
            )

        order_by = parsed.get("order_by")
        if order_by:
            order_col = str(order_by["column"])
            if order_col not in headers:
                raise ValueError(
                    f"Unknown column: {order_col}. Available columns: {headers}"
                )
            reverse = order_by["direction"] == "DESC"
            rows = sorted(
                rows,
                key=lambda r: self._sort_key(r.get(order_col)),
                reverse=reverse,
            )

        offset = parsed.get("offset") or 0
        limit = parsed.get("limit")
        if offset:
            rows = rows[offset:]
        if limit is not None:
            rows = rows[:limit]

        rows_out = [tuple(row.get(col) for col in selected_columns) for row in rows]

        distinct = parsed.get("distinct", False)
        if distinct:
            seen: set[tuple[Any, ...]] = set()
            unique_rows: list[tuple[Any, ...]] = []
            for r in rows_out:
                h = tuple(tuple(v) if isinstance(v, list) else v for v in r)
                if h not in seen:
                    seen.add(h)
                    unique_rows.append(r)
            rows_out = unique_rows

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

    def _execute_aggregate_select(
        self,
        action: str,
        parsed: dict[str, Any],
        headers: list[str],
        rows: list[dict[str, Any]],
        columns: list[Any],
        group_by: list[str] | None,
        having: dict[str, Any] | None,
    ) -> ExecutionResult:
        if group_by is not None:
            missing_group_columns = [col for col in group_by if col not in headers]
            if missing_group_columns:
                raise ValueError(
                    "Unknown column(s): "
                    f"{', '.join(missing_group_columns)}. Available columns: {headers}"
                )

        output_columns: list[str] = []
        for column in columns:
            if self._is_aggregate_column(column):
                arg = str(column.get("arg"))
                if arg != "*" and arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                output_columns.append(self._aggregate_label(column))
                continue

            column_name = str(column)
            if column_name not in headers:
                raise ValueError(
                    f"Unknown column(s): {column_name}. Available columns: {headers}"
                )
            if group_by is None:
                raise ValueError(
                    "Non-aggregate columns in aggregate queries require GROUP BY"
                )
            if column_name not in group_by:
                raise ValueError(
                    f"Selected column '{column_name}' must appear in GROUP BY"
                )
            output_columns.append(column_name)

        grouped_rows: list[dict[str, Any]] = []
        if group_by:
            groups: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
            for row in rows:
                group_key = tuple(row.get(col) for col in group_by)
                groups.setdefault(group_key, []).append(row)

            for group_key, group_values in groups.items():
                group_map = dict(zip(group_by, group_key))
                output_row: dict[str, Any] = {}
                for column in columns:
                    if self._is_aggregate_column(column):
                        label = self._aggregate_label(column)
                        output_row[label] = self._compute_aggregate(
                            str(column.get("func")),
                            str(column.get("arg")),
                            group_values,
                        )
                    else:
                        col_name = str(column)
                        output_row[col_name] = group_map.get(col_name)
                if having and not self._matches_where(output_row, having):
                    continue
                grouped_rows.append(output_row)
        else:
            output_row_single: dict[str, Any] = {}
            for column in columns:
                if self._is_aggregate_column(column):
                    label = self._aggregate_label(column)
                    output_row_single[label] = self._compute_aggregate(
                        str(column.get("func")),
                        str(column.get("arg")),
                        rows,
                    )
            if not having or self._matches_where(output_row_single, having):
                grouped_rows.append(output_row_single)

        distinct = parsed.get("distinct", False)
        if distinct:
            seen: set[tuple[Any, ...]] = set()
            deduped: list[dict[str, Any]] = []
            for row in grouped_rows:
                key = tuple(row.get(col) for col in output_columns)
                if key not in seen:
                    seen.add(key)
                    deduped.append(row)
            grouped_rows = deduped

        order_by = parsed.get("order_by")
        if order_by:
            order_column = str(order_by["column"])
            if order_column not in output_columns:
                raise ValueError(
                    f"Unknown column: {order_column}. Available columns: {output_columns}"
                )
            reverse = order_by["direction"] == "DESC"
            grouped_rows = sorted(
                grouped_rows,
                key=lambda r: self._sort_key(r.get(order_column)),
                reverse=reverse,
            )

        offset = parsed.get("offset") or 0
        limit = parsed.get("limit")
        if offset:
            grouped_rows = grouped_rows[offset:]
        if limit is not None:
            grouped_rows = grouped_rows[:limit]

        rows_out = [
            tuple(row.get(col) for col in output_columns) for row in grouped_rows
        ]
        description: Description = [
            (col, None, None, None, None, None, None) for col in output_columns
        ]
        return ExecutionResult(
            action=action,
            rows=rows_out,
            description=description,
            rowcount=len(rows_out),
            lastrowid=None,
        )

    def _is_aggregate_column(self, column: Any) -> bool:
        return (
            isinstance(column, dict)
            and column.get("type") == "aggregate"
            and isinstance(column.get("func"), str)
            and isinstance(column.get("arg"), str)
        )

    def _aggregate_label(self, aggregate: dict[str, Any]) -> str:
        return f"{str(aggregate['func']).upper()}({aggregate['arg']})"

    def _compute_aggregate(
        self, func: str, arg: str, rows: list[dict[str, Any]]
    ) -> Any:
        aggregate = func.upper()
        if aggregate == "COUNT":
            if arg == "*":
                return len(rows)
            return sum(1 for row in rows if row.get(arg) is not None)

        values = [row.get(arg) for row in rows if row.get(arg) is not None]
        if not values:
            return None

        numeric_values: list[float] = []
        for value in values:
            numeric = self._to_number(value)
            if numeric is not None:
                numeric_values.append(numeric)

        if not numeric_values:
            return None
        if aggregate == "SUM":
            return sum(numeric_values)
        if aggregate == "AVG":
            return sum(numeric_values) / len(numeric_values)
        if aggregate == "MIN":
            return min(numeric_values)
        if aggregate == "MAX":
            return max(numeric_values)
        raise ValueError(f"Unsupported aggregate function: {func}")

    def _evaluate_condition(
        self, row: dict[str, Any], condition: dict[str, Any]
    ) -> bool:
        column = condition["column"]
        operator = condition["operator"]
        value = condition["value"]
        row_value = row.get(column)

        if operator == "IS" and value is None:
            return row_value is None
        if operator == "IS NOT" and value is None:
            return row_value is not None
        if operator == "IN":
            if row_value is None:
                return False
            for candidate in value:
                left, right = self._coerce_for_compare(row_value, candidate)
                if left == right:
                    return True
            return False
        if operator == "BETWEEN":
            if row_value is None:
                return False
            low, high = value
            if low is None or high is None:
                return False
            left_low, right_low = self._coerce_for_compare(row_value, low)
            left_high, right_high = self._coerce_for_compare(row_value, high)
            return bool(left_low >= right_low and left_high <= right_high)
        if operator == "LIKE":
            if row_value is None:
                return False
            if not isinstance(value, str):
                raise NotImplementedError("Unsupported LIKE pattern type")
            regex = "^" + re.escape(value).replace(r"%", ".*").replace(r"_", ".") + "$"
            return bool(re.match(regex, str(row_value)))

        if row_value is None or value is None:
            return False

        left, right = self._coerce_for_compare(row_value, value)
        if operator in {"=", "=="}:
            return bool(left == right)
        if operator in {"!=", "<>"}:
            return bool(left != right)
        if operator == ">":
            return bool(left > right)
        if operator == ">=":
            return bool(left >= right)
        if operator == "<":
            return bool(left < right)
        if operator == "<=":
            return bool(left <= right)
        raise NotImplementedError(f"Unsupported operator: {operator}")

    def _coerce_for_compare(self, left: Any, right: Any) -> tuple[Any, Any]:
        left_num = self._to_number(left)
        right_num = self._to_number(right)
        if left_num is not None and right_num is not None:
            return left_num, right_num
        return str(left if left is not None else ""), str(
            right if right is not None else ""
        )

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

    def _resolve_sheet_name(self, requested_name: str) -> str | None:
        sheets = self.backend.list_sheets()
        lowered = {name.lower(): name for name in sheets}
        return lowered.get(requested_name.lower())
