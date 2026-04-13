import copy
import re
from typing import Any, Callable

from .engines.base import TableData, WorkbookBackend
from .engines.result import Description, ExecutionResult
from .exceptions import ProgrammingError
from .parser import _parse_column_expression, parse_sql
from .sanitize import sanitize_cell_value, sanitize_row

_READONLY_ACTIONS = frozenset(
    {"INSERT", "UPDATE", "DELETE", "CREATE", "DROP", "ALTER"}
)


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

        if action == "COMPOUND":
            return self._execute_compound(parsed)

        table = parsed["table"]
        resolved_table = self._resolve_sheet_name(table)

        if action == "SELECT":
            if resolved_table is None:
                available = self.backend.list_sheets()
                msg = f"Sheet '{table}' not found in Excel."
                if available:
                    msg += f" Available sheets: {available}"
                raise ValueError(msg)
            if parsed.get("joins") is not None:
                return self._execute_join_select(action, parsed, resolved_table)
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
            if where:
                where = copy.deepcopy(where)
                self._resolve_subqueries(where)
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
                    raw_value = update["value"]
                    if isinstance(raw_value, dict) and raw_value.get("type") in {"case", "binary_op", "unary_op", "literal"}:
                        evaluated = self._eval_expression(
                            raw_value, row_map, lambda c: row_map.get(c)
                        )
                    else:
                        evaluated = raw_value
                    value = (
                        sanitize_cell_value(evaluated)
                        if self.sanitize_formulas
                        else evaluated
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
            if where:
                where = copy.deepcopy(where)
                self._resolve_subqueries(where)
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

            if insert_columns is not None:
                missing = [col for col in insert_columns if col not in headers]
                if missing:
                    raise ValueError(
                        f"Unknown column(s): {', '.join(missing)}. Available columns: {headers}"
                    )

            rows_to_insert: list[list[Any]]
            if isinstance(values, dict):
                if values.get("type") != "subquery" or "query" not in values:
                    raise ValueError("Invalid INSERT subquery format")
                subquery_result = self.execute(values["query"])
                rows_to_insert = [list(row) for row in subquery_result.rows]
                # Validate column count from subquery even when zero rows returned
                expected_count = (
                    len(insert_columns) if insert_columns is not None else len(headers)
                )
                if subquery_result.description:
                    actual_count = len(subquery_result.description)
                    if actual_count != expected_count:
                        raise ValueError(
                            f"INSERT...SELECT column count mismatch: "
                            f"target has {expected_count} column(s), "
                            f"SELECT returns {actual_count}"
                        )
            elif isinstance(values, list):
                rows_to_insert = [list(row) for row in values]
            else:
                raise ValueError("Invalid INSERT values format")

            # Pre-validate ALL rows before appending any (atomicity guarantee)
            expected_count = (
                len(insert_columns) if insert_columns is not None else len(headers)
            )
            sanitized_rows: list[list[Any]] = []
            for values_row in rows_to_insert:
                if len(values_row) != expected_count:
                    if insert_columns is None:
                        raise ValueError("INSERT values count does not match header count")
                    else:
                        raise ValueError("INSERT values count does not match column count")
                if insert_columns is None:
                    row_values = list(values_row)
                else:
                    row_values = [None for _ in headers]
                    for col, value in zip(insert_columns, values_row):
                        row_values[headers.index(col)] = value
                sanitized_row = (
                    sanitize_row(row_values) if self.sanitize_formulas else row_values
                )
                sanitized_rows.append(sanitized_row)

            on_conflict = parsed.get("on_conflict")
            if on_conflict is not None:
                target_cols = on_conflict["target_columns"]
                for target_col in target_cols:
                    if target_col not in headers:
                        raise ValueError(f"ON CONFLICT column '{target_col}' not found in headers")

                target_indices = [headers.index(target_col) for target_col in target_cols]
                action_name = str(on_conflict.get("action", "")).upper()
                upsert_updates = on_conflict.get("set", [])
                if action_name == "UPDATE":
                    for update in upsert_updates:
                        if update["column"] not in headers:
                            raise ValueError(
                                f"Unknown column: {update['column']}. Available columns: {headers}"
                            )

                rowcount = 0
                for sanitized_row in sanitized_rows:
                    conflict_row: list[Any] | None = None
                    for existing_row in table_data.rows:
                        is_conflict = True
                        for target_index in target_indices:
                            existing_value = (
                                existing_row[target_index]
                                if target_index < len(existing_row)
                                else None
                            )
                            incoming_value = (
                                sanitized_row[target_index]
                                if target_index < len(sanitized_row)
                                else None
                            )
                            if existing_value is None or incoming_value is None:
                                # SQL semantics: NULL never matches NULL for conflict detection
                                is_conflict = False
                                break
                            left, right = self._coerce_for_compare(existing_value, incoming_value)
                            if left != right:
                                is_conflict = False
                                break
                        if is_conflict:
                            conflict_row = existing_row
                            break

                    if conflict_row is None:
                        table_data.rows.append(list(sanitized_row))
                        rowcount += 1
                        continue

                    if action_name == "NOTHING":
                        continue

                    if action_name != "UPDATE":
                        raise ValueError(
                            f"Invalid ON CONFLICT action: {on_conflict.get('action')}"
                        )

                    row_map = self._row_from_values(headers, conflict_row)
                    excluded_map = self._row_from_values(headers, sanitized_row)

                    def _resolve_upsert_column(column_name: str) -> Any:
                        if column_name.startswith("excluded."):
                            return excluded_map.get(column_name.split(".", 1)[1])
                        return row_map.get(column_name)

                    for update in upsert_updates:
                        col_index = headers.index(update["column"])
                        raw_value = update["value"]
                        if isinstance(raw_value, dict) and raw_value.get("type") in {
                            "alias",
                            "case",
                            "binary_op",
                            "unary_op",
                            "literal",
                            "column",
                        }:
                            evaluated = self._eval_expression(
                                raw_value,
                                row_map,
                                _resolve_upsert_column,
                            )
                        else:
                            evaluated = raw_value
                        value = (
                            sanitize_cell_value(evaluated)
                            if self.sanitize_formulas
                            else evaluated
                        )
                        if col_index >= len(conflict_row):
                            conflict_row.extend([None] * (col_index - len(conflict_row) + 1))
                        conflict_row[col_index] = value
                    rowcount += 1

                self.backend.write_sheet(resolved_table, table_data)
                return ExecutionResult(
                    action=action,
                    rows=[],
                    description=[],
                    rowcount=rowcount,
                    lastrowid=None,
                )

            # All rows validated — now append atomically
            last_row = None
            for sanitized_row in sanitized_rows:
                last_row = self.backend.append_row(resolved_table, sanitized_row)

            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=len(rows_to_insert),
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

        if action == "ALTER":
            if resolved_table is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            operation = parsed.get("operation")
            data = self.backend.read_sheet(resolved_table)

            if operation == "ADD_COLUMN":
                col = parsed["column"]
                if col in data.headers:
                    raise ValueError(f"Column '{col}' already exists in '{table}'")
                data.headers.append(col)
                for row in data.rows:
                    row.append(None)
                self.backend.write_sheet(resolved_table, data)
            elif operation == "DROP_COLUMN":
                col = parsed["column"]
                if col not in data.headers:
                    raise ValueError(f"Column '{col}' not found in '{table}'")
                if len(data.headers) == 1:
                    raise ValueError(
                        f"Cannot drop the only column '{col}' from '{table}'"
                    )
                idx = data.headers.index(col)
                data.headers.pop(idx)
                for row in data.rows:
                    if idx < len(row):
                        row.pop(idx)
                self.backend.write_sheet(resolved_table, data)
            elif operation == "RENAME_COLUMN":
                old_col = parsed["old_column"]
                new_col = parsed["new_column"]
                if old_col not in data.headers:
                    raise ValueError(f"Column '{old_col}' not found in '{table}'")
                if new_col in data.headers:
                    raise ValueError(f"Column '{new_col}' already exists in '{table}'")
                idx = data.headers.index(old_col)
                data.headers[idx] = new_col
                self.backend.write_sheet(resolved_table, data)
            else:
                raise ValueError(f"Unsupported ALTER operation: {operation}")

            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        raise ValueError(f"Unsupported action: {action}")

    def _normalize_row_key(self, row: tuple[Any, ...]) -> tuple[Any, ...]:
        return tuple(tuple(value) if isinstance(value, list) else value for value in row)

    def _dedupe_rows(self, rows: list[tuple[Any, ...]]) -> list[tuple[Any, ...]]:
        seen: set[tuple[Any, ...]] = set()
        deduped: list[tuple[Any, ...]] = []
        for row in rows:
            key = self._normalize_row_key(row)
            if key in seen:
                continue
            seen.add(key)
            deduped.append(row)
        return deduped

    @staticmethod
    def _normalize_order_by(
        raw: Any,
    ) -> list[dict[str, str]] | None:
        """Normalize order_by from parsed dict to a list of order items."""
        if isinstance(raw, list):
            return raw  # already a list
        if isinstance(raw, dict):
            return [raw]
        return None

    @staticmethod
    def _unwrap_alias(column: Any) -> Any:
        if isinstance(column, dict) and column.get("type") == "alias":
            return column["expression"]
        return column

    @staticmethod
    def _output_name(column: Any) -> str:
        if isinstance(column, dict):
            if column.get("type") == "alias":
                return str(column["alias"])
            if column.get("type") == "aggregate":
                return f"{column['func']}({column['arg']})"
            if column.get("type") == "column":
                return f"{column['source']}.{column['name']}"
            if column.get("type") in {"binary_op", "unary_op", "literal"}:
                return SharedExecutor._expression_label(column)
            if column.get("type") == "case":
                return "case_expr"
        return str(column)

    @staticmethod
    def _source_key(column: Any) -> str:
        inner = column
        if isinstance(column, dict) and column.get("type") == "alias":
            inner = column["expression"]
        if isinstance(inner, dict):
            if inner.get("type") == "aggregate":
                return f"{inner['func']}({inner['arg']})"
            if inner.get("type") == "column":
                return f"{inner['source']}.{inner['name']}"
            if inner.get("type") in {"binary_op", "unary_op", "literal", "case"}:
                return f"__expr__:{SharedExecutor._expression_to_sql(inner)}"
        return str(inner)

    @staticmethod
    def _expression_to_sql(expression: Any) -> str:
        if isinstance(expression, dict):
            expression_type = expression.get("type")
            if expression_type == "alias":
                return SharedExecutor._expression_to_sql(expression.get("expression"))
            if expression_type == "column":
                return f"{expression['source']}.{expression['name']}"
            if expression_type == "aggregate":
                return f"{expression['func']}({expression['arg']})"
            if expression_type == "literal":
                return SharedExecutor._literal_to_sql(expression.get("value"))
            if expression_type == "unary_op":
                operand_sql = SharedExecutor._expression_to_sql(expression.get("operand"))
                return f"-{operand_sql}"
            if expression_type == "binary_op":
                left_sql = SharedExecutor._expression_to_sql(expression.get("left"))
                right_sql = SharedExecutor._expression_to_sql(expression.get("right"))
                return f"({left_sql} {expression['op']} {right_sql})"
            if expression_type == "case":
                parts: list[str] = ["CASE"]
                mode = expression.get("mode", "searched")
                if mode == "simple" and expression.get("value") is not None:
                    parts.append(SharedExecutor._expression_to_sql(expression["value"]))
                for when_branch in expression.get("whens", []):
                    if mode == "searched":
                        condition = when_branch.get("condition")
                        condition_sql = ""
                        if isinstance(condition, dict):
                            condition_sql = SharedExecutor._where_to_sql(condition)
                        parts.append(f"WHEN {condition_sql} THEN")
                    else:
                        match_sql = SharedExecutor._expression_to_sql(when_branch.get("match"))
                        parts.append(f"WHEN {match_sql} THEN")
                    parts.append(SharedExecutor._expression_to_sql(when_branch["result"]))
                if expression.get("else") is not None:
                    parts.append("ELSE")
                    parts.append(SharedExecutor._expression_to_sql(expression["else"]))
                parts.append("END")
                return " ".join(parts)
        return str(expression)

    @staticmethod
    def _literal_to_sql(value: Any) -> str:
        if value is None:
            return "NULL"
        if isinstance(value, str):
            escaped = value.replace("'", "''")
            return f"'{escaped}'"
        return str(value)

    @staticmethod
    def _where_operand_to_sql(operand: Any, *, is_column: bool) -> str:
        if isinstance(operand, dict):
            if operand.get("type") == "subquery":
                return "(SUBQUERY)"
            return SharedExecutor._expression_to_sql(operand)
        if isinstance(operand, str):
            if is_column:
                return operand
            return SharedExecutor._literal_to_sql(operand)
        return SharedExecutor._literal_to_sql(operand)

    @staticmethod
    def _where_to_sql(where: dict[str, Any]) -> str:
        node_type = where.get("type")
        if node_type == "not":
            operand = where.get("operand")
            if isinstance(operand, dict):
                return f"NOT ({SharedExecutor._where_to_sql(operand)})"
            return "NOT"

        if "conditions" in where:
            conditions = where.get("conditions", [])
            if not conditions:
                return ""
            parts = [SharedExecutor._where_to_sql(conditions[0])]
            conjunctions = where.get("conjunctions", [])
            for idx, conjunction in enumerate(conjunctions):
                if idx + 1 >= len(conditions):
                    break
                parts.append(str(conjunction))
                parts.append(SharedExecutor._where_to_sql(conditions[idx + 1]))
            combined = " ".join(parts)
            if where.get("type") == "compound":
                return f"({combined})"
            return combined

        column_sql = SharedExecutor._where_operand_to_sql(
            where.get("column"),
            is_column=True,
        )
        operator = str(where.get("operator", ""))
        value = where.get("value")

        if operator in {"IS", "IS NOT"}:
            return f"{column_sql} {operator} NULL"

        if operator in {"IN", "NOT IN"}:
            if isinstance(value, dict) and value.get("type") == "subquery":
                return f"{column_sql} {operator} (SUBQUERY)"
            if isinstance(value, (list, tuple)):
                values_sql = ", ".join(
                    SharedExecutor._where_operand_to_sql(item, is_column=False)
                    for item in value
                )
            else:
                values_sql = SharedExecutor._where_operand_to_sql(value, is_column=False)
            return f"{column_sql} {operator} ({values_sql})"

        if operator in {"BETWEEN", "NOT BETWEEN"} and isinstance(value, (list, tuple)):
            if len(value) != 2:
                return f"{column_sql} {operator}"
            low_sql = SharedExecutor._where_operand_to_sql(value[0], is_column=False)
            high_sql = SharedExecutor._where_operand_to_sql(value[1], is_column=False)
            return f"{column_sql} {operator} {low_sql} AND {high_sql}"

        value_sql = SharedExecutor._where_operand_to_sql(value, is_column=False)
        return f"{column_sql} {operator} {value_sql}"

    @staticmethod
    def _expression_label(expression: Any) -> str:
        expression_sql = SharedExecutor._expression_to_sql(expression)
        if expression_sql.startswith("(") and expression_sql.endswith(")"):
            return expression_sql[1:-1]
        return expression_sql

    @staticmethod
    def _contains_arithmetic_expression(column: Any) -> bool:
        inner = SharedExecutor._unwrap_alias(column)
        if not isinstance(inner, dict):
            return False
        expression_type = inner.get("type")
        if expression_type in {"binary_op", "unary_op", "literal", "case"}:
            return True
        if expression_type == "alias":
            return SharedExecutor._contains_arithmetic_expression(inner.get("expression"))
        return False

    @staticmethod
    def _collect_expression_column_refs(expression: Any) -> set[str]:
        refs: set[str] = set()
        if isinstance(expression, str):
            refs.add(expression)
            return refs
        if not isinstance(expression, dict):
            return refs

        expression_type = expression.get("type")
        if expression_type == "alias":
            return SharedExecutor._collect_expression_column_refs(expression.get("expression"))
        if expression_type == "column":
            refs.add(f"{expression['source']}.{expression['name']}")
            return refs
        if expression_type == "binary_op":
            refs.update(SharedExecutor._collect_expression_column_refs(expression.get("left")))
            refs.update(SharedExecutor._collect_expression_column_refs(expression.get("right")))
            return refs
        if expression_type == "unary_op":
            refs.update(SharedExecutor._collect_expression_column_refs(expression.get("operand")))
            return refs
        if expression_type == "case":
            if expression.get("value") is not None:
                refs.update(SharedExecutor._collect_expression_column_refs(expression["value"]))
            for when_branch in expression.get("whens", []):
                # Collect from condition (searched mode) or match (simple mode)
                condition = when_branch.get("condition")
                if condition is not None:
                    refs.update(SharedExecutor._collect_where_column_refs(condition))
                match_expr = when_branch.get("match")
                if match_expr is not None:
                    refs.update(SharedExecutor._collect_expression_column_refs(match_expr))
                refs.update(SharedExecutor._collect_expression_column_refs(when_branch["result"]))
            if expression.get("else") is not None:
                refs.update(SharedExecutor._collect_expression_column_refs(expression["else"]))
            return refs
        return refs

    @staticmethod
    def _collect_where_column_refs(where: dict[str, Any]) -> set[str]:
        """Collect column references from a WHERE condition tree."""
        refs: set[str] = set()
        node_type = where.get("type")
        if node_type == "not":
            refs.update(SharedExecutor._collect_where_column_refs(where["operand"]))
            return refs
        if "conditions" in where:
            for cond in where["conditions"]:
                refs.update(SharedExecutor._collect_where_column_refs(cond))
            return refs
        # Atomic condition: {column, operator, value}
        column = where.get("column")
        if isinstance(column, str):
            refs.add(column)
        elif isinstance(column, dict):
            refs.update(SharedExecutor._collect_expression_column_refs(column))

        value = where.get("value")
        if isinstance(value, dict) and value.get("type") != "subquery":
            refs.update(SharedExecutor._collect_expression_column_refs(value))
        elif isinstance(value, (list, tuple)):
            for candidate in value:
                if isinstance(candidate, dict):
                    refs.update(SharedExecutor._collect_expression_column_refs(candidate))
        return refs
    @staticmethod
    def _build_alias_map(columns: list[Any]) -> dict[str, str]:
        alias_map: dict[str, str] = {}
        for column in columns:
            if isinstance(column, dict) and column.get("type") == "alias":
                alias_name = str(column["alias"])
                alias_map[alias_name] = SharedExecutor._source_key(column)
        return alias_map

    def _apply_order_by(
        self,
        rows: list[Any],
        order_by: list[dict[str, Any]] | None,
        *,
        value_getter: Callable[[Any, str], Any],
        available_columns: set[str] | None = None,
    ) -> list[Any]:
        if not order_by:
            return rows
        if available_columns is not None:
            for item in order_by:
                col = str(item["column"])
                if col not in available_columns:
                    raise ValueError(
                        f"Unknown column: {col}. Available columns: {sorted(available_columns)}"
                    )
        if len(rows) < 2:
            return rows
        for item in reversed(order_by):
            col = str(item["column"])
            reverse = item["direction"] == "DESC"
            rows = sorted(
                rows,
                key=lambda r: self._sort_key(value_getter(r, col)),
                reverse=reverse,
            )
        return rows

    def _materialize_order_expression_columns(
        self,
        rows: list[dict[str, Any]],
        order_by: list[dict[str, Any]] | None,
    ) -> set[str]:
        expression_columns: set[str] = set()
        if not order_by:
            return expression_columns

        parsed_expressions: dict[str, Any] = {}
        for item in order_by:
            col_ref = str(item["column"])
            if not col_ref.startswith("__expr__:"):
                continue
            expression_columns.add(col_ref)
            if col_ref in parsed_expressions:
                continue
            expression_sql = col_ref[len("__expr__:") :]
            parsed_expressions[col_ref] = _parse_column_expression(
                expression_sql,
                allow_wildcard=False,
                allow_aggregates=False,
            )

        if not parsed_expressions:
            return expression_columns

        for row in rows:
            for col_ref, expression in parsed_expressions.items():
                row[col_ref] = self._eval_expression(
                    expression,
                    row,
                    lambda col_name: row.get(col_name),
                )

        return expression_columns

    def _execute_compound(self, parsed: dict[str, Any]) -> ExecutionResult:
        queries = parsed.get("queries")
        if not isinstance(queries, list) or not queries:
            raise ValueError("COMPOUND query must contain at least one SELECT query")

        operators = parsed.get("operators")
        if not isinstance(operators, list):
            operator = parsed.get("operator")
            if not isinstance(operator, str):
                raise ValueError("COMPOUND query must include a valid operator")
            operators = [operator] * (len(queries) - 1)

        if len(operators) != len(queries) - 1:
            raise ValueError("Invalid COMPOUND query structure")

        results = [self.execute(query) for query in queries]
        first_result = results[0]
        expected_columns = len(first_result.description)
        for result in results[1:]:
            if len(result.description) != expected_columns:
                raise ValueError("Compound queries require matching column counts")

        rows = list(first_result.rows)
        for idx, operator in enumerate(operators, start=1):
            next_rows = list(results[idx].rows)
            normalized_operator = operator.upper()

            if normalized_operator == "UNION ALL":
                rows = rows + next_rows
                continue

            if normalized_operator == "UNION":
                rows = self._dedupe_rows(rows + next_rows)
                continue

            if normalized_operator == "INTERSECT":
                right_keys = {self._normalize_row_key(row) for row in next_rows}
                rows = [
                    row
                    for row in self._dedupe_rows(rows)
                    if self._normalize_row_key(row) in right_keys
                ]
                continue

            if normalized_operator == "EXCEPT":
                right_keys = {self._normalize_row_key(row) for row in next_rows}
                rows = [
                    row
                    for row in self._dedupe_rows(rows)
                    if self._normalize_row_key(row) not in right_keys
                ]
                continue

            raise ValueError(f"Unsupported compound operator: {operator}")

        # Apply compound-level ORDER BY / LIMIT / OFFSET.
        order_by = self._normalize_order_by(parsed.get("order_by"))
        if order_by:
            desc_names = [d[0] for d in first_result.description]
            resolved_indexes: dict[str, int] = {}
            for item in order_by:
                col_name = str(item["column"])
                col_index: int | None = None
                for i, dname in enumerate(desc_names):
                    if dname is not None and (dname == col_name or dname.endswith(f".{col_name}")):
                        col_index = i
                        break
                if col_index is None:
                    raise ValueError(f"ORDER BY column '{col_name}' not found in compound result")
                resolved_indexes[col_name] = col_index
            rows = self._apply_order_by(
                rows,
                order_by,
                value_getter=lambda r, col: r[resolved_indexes[col]],
            )

        compound_offset = parsed.get("offset") or 0
        compound_limit = parsed.get("limit")
        if compound_offset:
            rows = rows[compound_offset:]
        if compound_limit is not None:
            rows = rows[:compound_limit]

        return ExecutionResult(
            action="COMPOUND",
            rows=rows,
            description=first_result.description,
            rowcount=len(rows),
            lastrowid=None,
        )

    @staticmethod
    def _validate_join_where_refs(
        where: dict[str, Any],
        validate_fn: Callable[[str, str, str], None],
    ) -> None:
        """Recursively validate column references in WHERE tree for JOIN queries."""
        for condition in where.get("conditions", []):
            SharedExecutor._validate_join_where_node(condition, validate_fn)

    @staticmethod
    def _validate_join_where_node(
        node: dict[str, Any],
        validate_fn: Callable[[str, str, str], None],
    ) -> None:
        """Validate a single WHERE AST node for JOIN column refs."""
        # NOT node: recurse into operand
        if node.get("type") == "not":
            operand = node.get("operand")
            if isinstance(operand, dict):
                SharedExecutor._validate_join_where_node(operand, validate_fn)
            return
        # Compound or precedence-grouped node: recurse
        if "conditions" in node and node.get("type") != "not":
            for child in node["conditions"]:
                SharedExecutor._validate_join_where_node(child, validate_fn)
            return

        def _validate_expression_refs(expression: Any, *, column_context: bool) -> None:
            if expression is None:
                return
            if isinstance(expression, dict):
                expression_type = expression.get("type")
                if expression_type == "column":
                    validate_fn(
                        str(expression["source"]),
                        str(expression["name"]),
                        "WHERE",
                    )
                    return
                if expression_type == "alias":
                    _validate_expression_refs(expression.get("expression"), column_context=column_context)
                    return
                if expression_type == "unary_op":
                    _validate_expression_refs(expression.get("operand"), column_context=False)
                    return
                if expression_type == "binary_op":
                    _validate_expression_refs(expression.get("left"), column_context=False)
                    _validate_expression_refs(expression.get("right"), column_context=False)
                    return
                if expression_type == "case":
                    mode = str(expression.get("mode", ""))
                    if mode == "simple":
                        _validate_expression_refs(expression.get("value"), column_context=False)
                    for when_branch in expression.get("whens", []):
                        if not isinstance(when_branch, dict):
                            continue
                        if mode == "searched":
                            condition = when_branch.get("condition")
                            if isinstance(condition, dict):
                                SharedExecutor._validate_join_where_node(condition, validate_fn)
                        else:
                            _validate_expression_refs(when_branch.get("match"), column_context=False)
                        _validate_expression_refs(when_branch.get("result"), column_context=False)
                    _validate_expression_refs(expression.get("else"), column_context=False)
                    return
                return
            if column_context and isinstance(expression, str) and "." in expression:
                src, col_name = expression.split(".", 1)
                validate_fn(src, col_name, "WHERE")

        column_expr = node.get("column")
        _validate_expression_refs(column_expr, column_context=True)

        value_expr = node.get("value")
        if isinstance(value_expr, dict) and value_expr.get("type") != "subquery":
            _validate_expression_refs(value_expr, column_context=False)
        elif isinstance(value_expr, (list, tuple)):
            for candidate in value_expr:
                _validate_expression_refs(candidate, column_context=False)

    def _matches_where(self, row: dict[str, Any], where: dict[str, Any]) -> bool:
        node_type = where.get("type")
        if node_type == "not":
            return not self._matches_where(row, where["operand"])
        if "conditions" in where:
            conditions = where["conditions"]
            conjunctions = where["conjunctions"]
            results = [self._matches_where(row, conditions[0])]
            for idx, conj in enumerate(conjunctions):
                next_result = self._matches_where(row, conditions[idx + 1])
                if conj == "AND":
                    results[-1] = results[-1] and next_result
                else:
                    results.append(next_result)
            return any(results)

        return self._evaluate_condition(row, where)

    def _row_from_values(self, headers: list[str], row_values: list[Any]) -> dict[str, Any]:
        return {
            headers[col_index]: row_values[col_index]
            if col_index < len(row_values)
            else None
            for col_index in range(len(headers))
        }

    def _build_source_row(
        self,
        source: dict[str, Any],
        headers: list[str],
        row_values: list[Any],
    ) -> dict[str, dict[str, Any]]:
        row_map = self._row_from_values(headers, row_values)
        source_row = {str(source["table"]): row_map}
        ref = str(source["ref"])
        if ref != str(source["table"]):
            source_row[ref] = row_map
        return source_row

    def _resolve_join_column(self, row: dict[str, Any], col_spec: dict[str, Any]) -> Any:
        source = str(col_spec.get("source", ""))
        column = str(col_spec.get("name", ""))
        source_row = row.get(source)
        if not isinstance(source_row, dict):
            raise ValueError(f"Unknown source reference: {source}")
        return source_row.get(column)

    def _flatten_join_row(self, row: dict[str, Any]) -> dict[str, Any]:
        flattened: dict[str, Any] = {}
        for source, source_row in row.items():
            if not isinstance(source_row, dict):
                continue
            for column, value in source_row.items():
                flattened[f"{source}.{column}"] = value
        return flattened

    def _join_two_sources(
        self,
        left_rows: list[dict[str, dict[str, Any]]],
        left_headers_map: dict[str, set[str]],
        right_data: TableData,
        right_source: dict[str, Any],
        join_type: str,
        on_clauses: list[dict[str, Any]],
        all_known_sources: set[str],
    ) -> tuple[list[dict[str, dict[str, Any]]], dict[str, set[str]]]:
        right_headers = list(right_data.headers)
        right_sources = {str(right_source["table"]), str(right_source["ref"])}

        # Sentinel: each NULL key gets a unique object so NULL != NULL per SQL standard.
        _null_sentinel_counter = 0

        def _normalize_key_value(val: Any) -> Any:
            """Coerce key values for consistent hash matching."""
            if val is None:
                nonlocal _null_sentinel_counter
                _null_sentinel_counter += 1
                return object()
            num = self._to_number(val)
            if num is not None:
                return num
            return val

        def build_key(
            row_ns: dict[str, Any],
            is_left_side: bool,
        ) -> tuple[Any, ...]:
            key_parts: list[Any] = []
            for clause in on_clauses:
                left_col = clause["left"]
                right_col = clause["right"]
                left_is_left = str(left_col["source"]) in all_known_sources
                selected = left_col if left_is_left == is_left_side else right_col
                raw_val = self._resolve_join_column(row_ns, selected)
                key_parts.append(_normalize_key_value(raw_val))
            return tuple(key_parts)

        right_rows = [
            self._build_source_row(right_source, right_headers, right_row_values)
            for right_row_values in right_data.rows
        ]
        join_type_upper = join_type.upper()
        right_hash: dict[tuple[Any, ...], list[dict[str, dict[str, Any]]]] = {}
        if join_type_upper != "CROSS":
            for right_ns in right_rows:
                key = build_key(right_ns, False)
                right_hash.setdefault(key, []).append(right_ns)

        joined_rows: list[dict[str, dict[str, Any]]] = []
        right_null_values = [None for _ in right_headers]
        right_null_ns = self._build_source_row(right_source, right_headers, right_null_values)

        if join_type_upper == "RIGHT":
            left_hash: dict[tuple[Any, ...], list[dict[str, dict[str, Any]]]] = {}
            for left_ns in left_rows:
                key = build_key(left_ns, True)
                left_hash.setdefault(key, []).append(left_ns)

            left_null_ns = {
                source: {column: None for column in columns}
                for source, columns in left_headers_map.items()
            }

            for right_ns in right_rows:
                key = build_key(right_ns, False)
                matches = left_hash.get(key, [])
                if matches:
                    for left_ns in matches:
                        combined_row = {}
                        combined_row.update(left_ns)
                        combined_row.update(right_ns)
                        joined_rows.append(combined_row)
                else:
                    combined_row = {}
                    combined_row.update(left_null_ns)
                    combined_row.update(right_ns)
                    joined_rows.append(combined_row)
        elif join_type_upper == "FULL":
            left_null_ns = {
                source: {column: None for column in columns}
                for source, columns in left_headers_map.items()
            }

            matched_right_indices: set[int] = set()

            for left_ns in left_rows:
                key = build_key(left_ns, True)
                matches = right_hash.get(key, [])
                if matches:
                    for right_ns in matches:
                        matched_right_indices.add(id(right_ns))
                        combined_row = {}
                        combined_row.update(left_ns)
                        combined_row.update(right_ns)
                        joined_rows.append(combined_row)
                else:
                    combined_row = {}
                    combined_row.update(left_ns)
                    combined_row.update(right_null_ns)
                    joined_rows.append(combined_row)

            for right_ns in right_rows:
                if id(right_ns) not in matched_right_indices:
                    combined_row = {}
                    combined_row.update(left_null_ns)
                    combined_row.update(right_ns)
                    joined_rows.append(combined_row)
        elif join_type_upper == "CROSS":
            for left_ns in left_rows:
                for right_ns in right_rows:
                    combined_row = {}
                    combined_row.update(left_ns)
                    combined_row.update(right_ns)
                    joined_rows.append(combined_row)
        else:
            for left_ns in left_rows:
                key = build_key(left_ns, True)
                matches = right_hash.get(key, [])
                if matches:
                    for right_ns in matches:
                        combined_row = {}
                        combined_row.update(left_ns)
                        combined_row.update(right_ns)
                        joined_rows.append(combined_row)
                elif join_type_upper == "LEFT":
                    combined_row = {}
                    combined_row.update(left_ns)
                    combined_row.update(right_null_ns)
                    joined_rows.append(combined_row)

        updated_headers_map = dict(left_headers_map)
        for source in right_sources:
            updated_headers_map[source] = set(right_headers)

        return joined_rows, updated_headers_map

    def _execute_join_select(
        self,
        action: str,
        parsed: dict[str, Any],
        resolved_left_table: str,
    ) -> ExecutionResult:
        joins = parsed.get("joins") or []
        from_source = parsed["from"]
        left_data = self.backend.read_sheet(resolved_left_table)
        if not left_data.headers:
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        left_headers = list(left_data.headers)

        source_headers: dict[str, set[str]] = {
            str(from_source["ref"]): set(left_headers),
            str(from_source["table"]): set(left_headers),
        }
        source_headers_ordered: list[tuple[str, list[str]]] = [
            (str(from_source["ref"]), left_headers),
        ]
        ref_to_table: dict[str, str] = {
            str(from_source["ref"]): str(from_source["table"]),
        }
        join_inputs: list[tuple[dict[str, Any], TableData]] = []
        known_sources = {str(from_source["table"]), str(from_source["ref"])}

        # --- Column existence validation ---
        def _validate_column_ref(source: str, name: str, context: str) -> None:
            valid = source_headers.get(source)
            if valid is None:
                raise ValueError(f"Unknown source reference: {source}")
            if name not in valid:
                raise ValueError(
                    f"Unknown column: {source}.{name}. "
                    f"Available columns in '{source}': {sorted(valid)}"
                )

        for join_spec in joins:
            right_source = join_spec["source"]
            resolved_right_table = self._resolve_sheet_name(str(right_source["table"]))
            if resolved_right_table is None:
                available = self.backend.list_sheets()
                msg = f"Sheet '{right_source['table']}' not found in Excel."
                if available:
                    msg += f" Available sheets: {available}"
                raise ValueError(msg)

            right_data = self.backend.read_sheet(resolved_right_table)
            if not right_data.headers:
                return ExecutionResult(
                    action=action,
                    rows=[],
                    description=[],
                    rowcount=0,
                    lastrowid=None,
                )

            right_headers = set(right_data.headers)
            right_ref = str(right_source["ref"])
            right_table_name = str(right_source["table"])
            source_headers[right_ref] = right_headers
            source_headers[right_table_name] = right_headers
            source_headers_ordered.append((right_ref, list(right_data.headers)))
            ref_to_table[right_ref] = right_table_name

            join_on = join_spec.get("on")
            on_clauses = (
                join_on.get("clauses", [])
                if join_on is not None
                else []
            )
            for clause in on_clauses:
                for side in ("left", "right"):
                    col = clause[side]
                    _validate_column_ref(str(col["source"]), str(col["name"]), "ON")

            known_sources.update({right_ref, right_table_name})
            join_inputs.append((join_spec, right_data))

        # Validate SELECT columns
        def _validate_join_select_expression(expression: Any) -> None:
            if expression is None:
                return
            if isinstance(expression, dict):
                expression_type = expression.get("type")
                if expression_type == "column":
                    _validate_column_ref(
                        str(expression["source"]),
                        str(expression["name"]),
                        "SELECT",
                    )
                    return
                if expression_type == "aggregate":
                    arg = str(expression.get("arg", ""))
                    if arg == "*":
                        return
                    if "." not in arg:
                        raise ValueError(
                            "Aggregate arguments in JOIN queries must be qualified column names or *"
                        )
                    source, name = arg.split(".", 1)
                    _validate_column_ref(source, name, "SELECT")
                    return
                if expression_type == "literal":
                    return
                if expression_type == "unary_op":
                    _validate_join_select_expression(expression.get("operand"))
                    return
                if expression_type == "binary_op":
                    _validate_join_select_expression(expression.get("left"))
                    _validate_join_select_expression(expression.get("right"))
                    return
                if expression_type == "case":
                    mode = str(expression.get("mode", ""))
                    if mode == "simple":
                        _validate_join_select_expression(expression.get("value"))
                    for when_branch in expression.get("whens", []):
                        if not isinstance(when_branch, dict):
                            continue
                        if mode == "searched":
                            condition = when_branch.get("condition")
                            if isinstance(condition, dict):
                                SharedExecutor._validate_join_where_node(condition, _validate_column_ref)
                        else:
                            _validate_join_select_expression(when_branch.get("match"))
                        _validate_join_select_expression(when_branch.get("result"))
                    _validate_join_select_expression(expression.get("else"))
                    return
            raise ValueError("JOIN queries require qualified column names in SELECT")

        if parsed["columns"] != ["*"]:
            for column in parsed["columns"]:
                inner = self._unwrap_alias(column)
                _validate_join_select_expression(inner)

        # Validate WHERE columns
        where_raw = parsed.get("where")
        if where_raw:
            self._validate_join_where_refs(where_raw, _validate_column_ref)

        alias_map = self._build_alias_map(parsed["columns"])
        order_by = self._normalize_order_by(parsed.get("order_by"))
        if order_by and alias_map:
            resolved_order_by: list[dict[str, Any]] = []
            for item in order_by:
                resolved_item = dict(item)
                column_name = str(item["column"])
                resolved_item["column"] = alias_map.get(column_name, column_name)
                resolved_order_by.append(resolved_item)
            order_by = resolved_order_by

        # Validate ORDER BY column
        if order_by:
            for item in order_by:
                col_ref = str(item["column"])
                if col_ref.startswith("__expr__:"):
                    continue
                if self._aggregate_spec_from_label(col_ref) is not None:
                    continue
                if "." in col_ref:
                    src, col_name = col_ref.split(".", 1)
                    _validate_column_ref(src, col_name, "ORDER BY")

        left_rows = [
            self._build_source_row(from_source, left_headers, left_row_values)
            for left_row_values in left_data.rows
        ]
        left_headers_map: dict[str, set[str]] = {
            str(from_source["table"]): set(left_headers),
            str(from_source["ref"]): set(left_headers),
        }
        all_known_sources = {str(from_source["table"]), str(from_source["ref"])}

        for join_spec, right_data in join_inputs:
            right_source = join_spec["source"]
            join_on = join_spec.get("on")
            join_on_clauses = (
                join_on.get("clauses", [])
                if join_on is not None
                else []
            )
            left_rows, left_headers_map = self._join_two_sources(
                left_rows=left_rows,
                left_headers_map=left_headers_map,
                right_data=right_data,
                right_source=right_source,
                join_type=str(join_spec["type"]),
                on_clauses=join_on_clauses,
                all_known_sources=all_known_sources,
            )
            all_known_sources.update(
                {
                    str(right_source["table"]),
                    str(right_source["ref"]),
                }
            )

        joined_rows_flat = [self._flatten_join_row(row) for row in left_rows]

        where = parsed.get("where")
        if where:
            joined_rows_flat = [
                row for row in joined_rows_flat if self._matches_where(row, where)
            ]

        columns = parsed["columns"]
        aggregate_query = (
            any(self._is_aggregate_column(self._unwrap_alias(col)) for col in columns)
            if columns != ["*"]
            else False
        )
        group_by = parsed.get("group_by")
        having = parsed.get("having")

        if columns == ["*"] and (aggregate_query or group_by is not None):
            raise ValueError("SELECT * is not supported with GROUP BY or aggregate functions")

        if aggregate_query or group_by is not None:
            flattened_headers: list[str] = []
            for source_ref, ordered_headers in source_headers_ordered:
                for col_name in ordered_headers:
                    flattened_headers.append(f"{source_ref}.{col_name}")
                table_name = ref_to_table.get(source_ref, source_ref)
                if table_name != source_ref:
                    for col_name in ordered_headers:
                        flattened_headers.append(f"{table_name}.{col_name}")

            return self._execute_aggregate_select(
                action,
                parsed,
                flattened_headers,
                joined_rows_flat,
                columns,
                group_by,
                having,
            )

        selected_columns: list[str] = []
        output_names: list[str] = []
        if columns == ["*"]:
            for source_ref, ordered_headers in source_headers_ordered:
                for col_name in ordered_headers:
                    selected_columns.append(f"{source_ref}.{col_name}")
                    output_names.append(f"{source_ref}.{col_name}")
            if order_by:
                order_expression_columns = self._materialize_order_expression_columns(
                    joined_rows_flat,
                    order_by,
                )
                available_cols: set[str] = set()
                for source_ref, ordered_headers in source_headers_ordered:
                    for col_name in ordered_headers:
                        available_cols.add(f"{source_ref}.{col_name}")
                available_cols.update(order_expression_columns)
                joined_rows_flat = self._apply_order_by(
                    joined_rows_flat,
                    order_by,
                    value_getter=lambda r, col: r.get(col),
                    available_columns=available_cols,
                )

            offset = parsed.get("offset") or 0
            limit = parsed.get("limit")
            if offset:
                joined_rows_flat = joined_rows_flat[offset:]
            if limit is not None:
                joined_rows_flat = joined_rows_flat[:limit]

            rows_out = [
                tuple(row.get(column_name) for column_name in selected_columns)
                for row in joined_rows_flat
            ]
        else:
            selected_columns = [self._source_key(column) for column in columns]
            output_names = [self._output_name(column) for column in columns]

            projected_rows: list[dict[str, Any]] = []
            for row in joined_rows_flat:
                projected_row = dict(row)
                for column, key in zip(columns, selected_columns):
                    inner = self._unwrap_alias(column)
                    projected_row[key] = self._eval_expression(
                        inner,
                        row,
                        lambda col_name: row.get(col_name),
                    )
                projected_rows.append(projected_row)

            if order_by:
                order_expression_columns = self._materialize_order_expression_columns(
                    projected_rows,
                    order_by,
                )
                available_columns = set(selected_columns)
                for source_ref, ordered_headers in source_headers_ordered:
                    for col_name in ordered_headers:
                        available_columns.add(f"{source_ref}.{col_name}")
                available_columns.update(order_expression_columns)
                projected_rows = self._apply_order_by(
                    projected_rows,
                    order_by,
                    value_getter=lambda r, col: r.get(col),
                    available_columns=available_columns,
                )

            offset = parsed.get("offset") or 0
            limit = parsed.get("limit")
            if offset:
                projected_rows = projected_rows[offset:]
            if limit is not None:
                projected_rows = projected_rows[:limit]

            rows_out = [
                tuple(row.get(column_name) for column_name in selected_columns)
                for row in projected_rows
            ]
        description: Description = [
            (column_name, None, None, None, None, None, None)
            for column_name in output_names
        ]
        return ExecutionResult(
            action=action,
            rows=rows_out,
            description=description,
            rowcount=len(rows_out),
            lastrowid=None,
        )

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
            where = copy.deepcopy(where)
            self._resolve_subqueries(where)
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
        selected_columns: list[str]
        output_names: list[str]
        needs_expression_projection = any(
            self._contains_arithmetic_expression(column) for column in columns
        )
        if columns == ["*"]:
            selected_columns = list(headers)
            output_names = list(headers)
        else:
            selected_columns = [self._source_key(column) for column in columns]
            output_names = [self._output_name(column) for column in columns]

        missing: list[str]
        if columns == ["*"]:
            missing = [col for col in selected_columns if col not in headers]
        elif needs_expression_projection:
            referenced_columns: set[str] = set()
            for column in columns:
                inner = self._unwrap_alias(column)
                referenced_columns.update(self._collect_expression_column_refs(inner))
            missing = sorted(col for col in referenced_columns if col not in headers)
        else:
            missing = [col for col in selected_columns if col not in headers]
        if missing:
            raise ValueError(
                f"Unknown column(s): {', '.join(missing)}. Available columns: {headers}"
            )

        alias_map = self._build_alias_map(columns)
        order_by = self._normalize_order_by(parsed.get("order_by"))
        if order_by and alias_map:
            resolved_order_by: list[dict[str, Any]] = []
            for item in order_by:
                resolved_item = dict(item)
                column_name = str(item["column"])
                resolved_item["column"] = alias_map.get(column_name, column_name)
                resolved_order_by.append(resolved_item)
            order_by = resolved_order_by

        order_expression_columns: set[str] = set()
        if order_by:
            order_expression_columns = self._materialize_order_expression_columns(
                rows,
                order_by,
            )

        if order_by:
            if needs_expression_projection and columns != ["*"]:
                pass
            else:
                available_columns = set(headers)
                available_columns.update(order_expression_columns)
                rows = self._apply_order_by(
                    rows,
                    order_by,
                    value_getter=lambda r, col: r.get(col),
                    available_columns=available_columns,
                )

        if needs_expression_projection and columns != ["*"]:
            projected_rows: list[dict[str, Any]] = []
            for row in rows:
                projected_row = dict(row)
                for column, key in zip(columns, selected_columns):
                    inner = self._unwrap_alias(column)
                    projected_row[key] = self._eval_expression(
                        inner,
                        row,
                        lambda col_name: row.get(col_name),
                    )
                projected_rows.append(projected_row)

            if order_by:
                expression_columns = self._materialize_order_expression_columns(
                    projected_rows,
                    order_by,
                )
                available_columns = set(headers)
                available_columns.update(selected_columns)
                available_columns.update(expression_columns)
                projected_rows = self._apply_order_by(
                    projected_rows,
                    order_by,
                    value_getter=lambda r, col: r.get(col),
                    available_columns=available_columns,
                )

            offset = parsed.get("offset") or 0
            limit = parsed.get("limit")
            if offset:
                projected_rows = projected_rows[offset:]
            if limit is not None:
                projected_rows = projected_rows[:limit]

            rows_out = [
                tuple(projected_row.get(col) for col in selected_columns)
                for projected_row in projected_rows
            ]
        else:
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
            (col, None, None, None, None, None, None) for col in output_names
        ]
        return ExecutionResult(
            action=action,
            rows=rows_out,
            description=description,
            rowcount=len(rows_out),
            lastrowid=None,
        )

    def _resolve_subqueries(self, where: dict[str, Any]) -> None:
        """Recursively resolve subqueries in the WHERE tree."""
        for condition in where.get("conditions", []):
            self._resolve_subquery_node(condition)

    def _resolve_subquery_node(self, node: dict[str, Any]) -> None:
        """Resolve subqueries in a single WHERE AST node."""
        # NOT node: recurse into operand
        if node.get("type") == "not":
            operand = node.get("operand")
            if isinstance(operand, dict):
                self._resolve_subquery_node(operand)
            return
        # Compound or precedence-grouped node: recurse into conditions
        if "conditions" in node and node.get("type") != "not":
            for child in node["conditions"]:
                self._resolve_subquery_node(child)
            return
        # Atomic condition: resolve subquery value
        value = node.get("value")
        if isinstance(value, dict) and value.get("type") == "subquery":
            subquery_result = self.execute(value["query"])
            node["value"] = tuple(row[0] for row in subquery_result.rows if row)

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
        output_sources: list[str] = []
        required_aggregates: dict[str, tuple[str, str, bool]] = {}
        for column in columns:
            inner = self._unwrap_alias(column)
            if self._is_aggregate_column(inner):
                aggregate_column = inner
                arg = str(aggregate_column.get("arg"))
                if arg != "*" and arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                label = self._aggregate_label(aggregate_column)
                required_aggregates[label] = (
                    str(aggregate_column.get("func")),
                    arg,
                    bool(aggregate_column.get("distinct")),
                )
                output_columns.append(self._output_name(column))
                output_sources.append(label)
                continue

            column_name = self._source_key(column)
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
            output_columns.append(self._output_name(column))
            output_sources.append(column_name)

        if having is not None:
            if "conditions" in having:
                refs = [str(condition["column"]) for condition in having["conditions"]]
            else:
                refs = [str(having["column"])]
            for ref in refs:
                aggregate_spec = self._aggregate_spec_from_label(ref)
                if aggregate_spec is None:
                    continue
                func, arg, distinct = aggregate_spec
                if arg != "*" and arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                required_aggregates[ref] = (func, arg, distinct)

        alias_map = self._build_alias_map(columns)
        order_by_clause = self._normalize_order_by(parsed.get("order_by"))
        if order_by_clause and alias_map:
            resolved_order_by: list[dict[str, Any]] = []
            for item in order_by_clause:
                resolved_item = dict(item)
                column_name = str(item["column"])
                resolved_item["column"] = alias_map.get(column_name, column_name)
                resolved_order_by.append(resolved_item)
            order_by_clause = resolved_order_by

        if order_by_clause is not None:
            for item in order_by_clause:
                ref = str(item["column"])
                aggregate_spec = self._aggregate_spec_from_label(ref)
                if aggregate_spec is None:
                    continue
                func, arg, distinct = aggregate_spec
                if arg != "*" and arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                required_aggregates[ref] = (func, arg, distinct)

        if having is not None:
            for condition in having["conditions"]:
                column = str(condition["column"])
                if self._aggregate_spec_from_label(column) is not None:
                    continue
                if group_by is None or column not in group_by:
                    raise ValueError(
                        f"Column '{column}' in HAVING must be a GROUP BY column or aggregate function"
                    )

        grouped_rows: list[dict[str, Any]] = []
        if group_by:
            groups: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
            for row in rows:
                group_key = tuple(row.get(col) for col in group_by)
                groups.setdefault(group_key, []).append(row)

            for group_key, group_values in groups.items():
                group_map = dict(zip(group_by, group_key))
                context_row: dict[str, Any] = dict(group_map)
                for label, (func, arg, distinct) in required_aggregates.items():
                    context_row[label] = self._compute_aggregate(
                        func,
                        arg,
                        group_values,
                        distinct=distinct,
                    )
                if having and not self._matches_where(context_row, having):
                    continue
                grouped_rows.append(context_row)
        else:
            context_row_single: dict[str, Any] = {}
            for label, (func, arg, distinct) in required_aggregates.items():
                context_row_single[label] = self._compute_aggregate(
                    func,
                    arg,
                    rows,
                    distinct=distinct,
                )
            if not having or self._matches_where(context_row_single, having):
                grouped_rows.append(context_row_single)

        distinct = parsed.get("distinct", False)
        if distinct:
            seen: set[tuple[Any, ...]] = set()
            deduped: list[dict[str, Any]] = []
            for row in grouped_rows:
                key = tuple(row.get(col) for col in output_sources)
                if key not in seen:
                    seen.add(key)
                    deduped.append(row)
            grouped_rows = deduped

        if order_by_clause:
            order_expression_columns = self._materialize_order_expression_columns(
                grouped_rows,
                order_by_clause,
            )
            available_order_columns = set(output_sources)
            if group_by:
                available_order_columns.update(group_by)
            available_order_columns.update(required_aggregates.keys())
            available_order_columns.update(order_expression_columns)
            grouped_rows = self._apply_order_by(
                grouped_rows,
                order_by_clause,
                value_getter=lambda r, col: r.get(col),
                available_columns=available_order_columns,
            )

        offset = parsed.get("offset") or 0
        limit = parsed.get("limit")
        if offset:
            grouped_rows = grouped_rows[offset:]
        if limit is not None:
            grouped_rows = grouped_rows[:limit]

        rows_out = [
            tuple(row.get(col) for col in output_sources) for row in grouped_rows
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
        column = self._unwrap_alias(column)
        return (
            isinstance(column, dict)
            and column.get("type") == "aggregate"
            and isinstance(column.get("func"), str)
            and isinstance(column.get("arg"), str)
        )

    def _aggregate_label(self, aggregate: dict[str, Any]) -> str:
        func = str(aggregate["func"]).upper()
        arg = str(aggregate["arg"])
        if aggregate.get("distinct"):
            return f"{func}(DISTINCT {arg})"
        return f"{func}({arg})"

    def _aggregate_spec_from_label(
        self, expression: str
    ) -> tuple[str, str, bool] | None:
        match = re.fullmatch(
            r"(?i)(COUNT|SUM|AVG|MIN|MAX)\((DISTINCT\s+)?([^\)]+)\)",
            expression.strip(),
        )
        if not match:
            return None
        func = match.group(1).upper()
        distinct_modifier = match.group(2)
        arg = match.group(3).strip()
        if not arg:
            return None
        if distinct_modifier and func != "COUNT":
            raise ValueError("DISTINCT is only supported with COUNT")
        if arg != "*" and not re.fullmatch(
            r"[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?", arg
        ):
            raise ValueError(
                f"Unsupported aggregate expression: {func}({arg}). "
                "Only bare or qualified column names and * are supported"
            )
        return func, arg, bool(distinct_modifier)

    def _compute_aggregate(
        self,
        func: str,
        arg: str,
        rows: list[dict[str, Any]],
        distinct: bool = False,
    ) -> Any:
        aggregate = func.upper()
        if aggregate == "COUNT":
            if distinct:
                if arg == "*":
                    raise ValueError("COUNT(DISTINCT *) is not supported")
                return len({row.get(arg) for row in rows if row.get(arg) is not None})
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

    def _eval_expression(
        self,
        expr: Any,
        row: dict[str, Any],
        resolve_column: Callable[[str], Any],
    ) -> Any:
        if isinstance(expr, str):
            return resolve_column(expr)

        if not isinstance(expr, dict):
            return expr

        expr_type = expr.get("type")
        if expr_type == "alias":
            return self._eval_expression(expr.get("expression"), row, resolve_column)

        if expr_type == "column":
            source = expr.get("source", expr.get("table"))
            if source is None:
                raise ProgrammingError("Column expression is missing source/table")
            column_key = f"{source}.{expr['name']}"
            return resolve_column(column_key)

        if expr_type == "literal":
            return expr.get("value")

        if expr_type == "unary_op":
            if expr.get("op") != "-":
                raise ProgrammingError(f"Unsupported unary operator: {expr.get('op')}")
            operand_value = self._eval_expression(expr.get("operand"), row, resolve_column)
            if operand_value is None:
                return None
            numeric_operand = self._to_number(operand_value)
            if numeric_operand is None:
                raise ProgrammingError("Arithmetic expression requires numeric operands")
            return -numeric_operand

        if expr_type == "binary_op":
            operator = str(expr.get("op"))
            left_value = self._eval_expression(expr.get("left"), row, resolve_column)
            right_value = self._eval_expression(expr.get("right"), row, resolve_column)
            if left_value is None or right_value is None:
                return None

            left_number = self._to_number(left_value)
            right_number = self._to_number(right_value)
            if left_number is None or right_number is None:
                raise ProgrammingError("Arithmetic expression requires numeric operands")

            if operator == "+":
                return left_number + right_number
            if operator == "-":
                return left_number - right_number
            if operator == "*":
                return left_number * right_number
            if operator == "/":
                if right_number == 0:
                    raise ProgrammingError("Division by zero in arithmetic expression")
                return left_number / right_number
            raise ProgrammingError(f"Unsupported arithmetic operator: {operator}")

        if expr_type == "aggregate":
            raise ProgrammingError("Aggregate expressions are not supported in row-level arithmetic")

        if expr_type == "case":
            mode = expr.get("mode", "searched")
            if mode == "searched":
                for when_branch in expr.get("whens", []):
                    if not isinstance(when_branch, dict):
                        continue
                    condition = when_branch.get("condition")
                    if isinstance(condition, dict) and self._matches_where(row, condition):
                        return self._eval_expression(when_branch["result"], row, resolve_column)
            else:  # simple mode
                case_value = self._eval_expression(expr.get("value"), row, resolve_column)
                for when_branch in expr.get("whens", []):
                    if not isinstance(when_branch, dict):
                        continue
                    match_value = self._eval_expression(when_branch["match"], row, resolve_column)
                    if case_value is not None and match_value is not None:
                        left, right = self._coerce_for_compare(case_value, match_value)
                        if left == right:
                            return self._eval_expression(when_branch["result"], row, resolve_column)
            # No WHEN matched — evaluate ELSE or return None
            else_expr = expr.get("else")
            if else_expr is not None:
                return self._eval_expression(else_expr, row, resolve_column)
            return None

        raise ProgrammingError(f"Unsupported expression type: {expr_type}")

    def _evaluate_condition(
        self, row: dict[str, Any], condition: dict[str, Any]
    ) -> bool:
        column = condition["column"]
        operator = condition["operator"]
        value = condition["value"]

        def _resolve_operand(operand: Any, *, as_column: bool) -> Any:
            if isinstance(operand, dict) and operand.get("type") != "subquery":
                return self._eval_expression(operand, row, lambda col_name: row.get(col_name))
            if as_column:
                return row.get(str(operand))
            return operand

        row_value = _resolve_operand(column, as_column=True)

        if operator == "IS" and value is None:
            return row_value is None
        if operator == "IS NOT" and value is None:
            return row_value is not None
        if operator == "IN":
            if row_value is None:
                return False
            for candidate in value:
                candidate_value = _resolve_operand(candidate, as_column=False)
                left, right = self._coerce_for_compare(row_value, candidate_value)
                if left == right:
                    return True
            return False
        if operator == "NOT IN":
            if row_value is None:
                return False
            for candidate in value:
                candidate_value = _resolve_operand(candidate, as_column=False)
                left, right = self._coerce_for_compare(row_value, candidate_value)
                if left == right:
                    return False
            return True
        if operator == "BETWEEN":
            if row_value is None:
                return False
            low, high = value
            low = _resolve_operand(low, as_column=False)
            high = _resolve_operand(high, as_column=False)
            if low is None or high is None:
                return False
            left_low, right_low = self._coerce_for_compare(row_value, low)
            left_high, right_high = self._coerce_for_compare(row_value, high)
            return bool(left_low >= right_low and left_high <= right_high)
        if operator == "NOT BETWEEN":
            if row_value is None:
                return False
            low, high = value
            low = _resolve_operand(low, as_column=False)
            high = _resolve_operand(high, as_column=False)
            if low is None or high is None:
                return False
            left_low, right_low = self._coerce_for_compare(row_value, low)
            left_high, right_high = self._coerce_for_compare(row_value, high)
            return not bool(left_low >= right_low and left_high <= right_high)
        if operator == "LIKE":
            if row_value is None:
                return False
            value = _resolve_operand(value, as_column=False)
            if not isinstance(value, str):
                raise NotImplementedError("Unsupported LIKE pattern type")
            regex = "^" + re.escape(value).replace(r"%", ".*").replace(r"_", ".") + "$"
            return bool(re.match(regex, str(row_value)))
        if operator == "NOT LIKE":
            if row_value is None:
                return False
            value = _resolve_operand(value, as_column=False)
            if not isinstance(value, str):
                raise NotImplementedError("Unsupported LIKE pattern type")
            regex = "^" + re.escape(value).replace(r"%", ".*").replace(r"_", ".") + "$"
            return not bool(re.match(regex, str(row_value)))

        value = _resolve_operand(value, as_column=False)

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
