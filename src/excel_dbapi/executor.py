import copy
import re
from typing import Any, Callable, Optional

from .engines.base import TableData, WorkbookBackend
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

    def _apply_order_by(
        self,
        rows: list[Any],
        order_by: list[dict[str, str]] | None,
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
        # Atomic condition: validate column reference
        col_ref = str(node.get("column", ""))
        if "." in col_ref:
            src, col_name = col_ref.split(".", 1)
            validate_fn(src, col_name, "WHERE")

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
        right_hash: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
        for right_ns in right_rows:
            key = build_key(right_ns, False)
            right_hash.setdefault(key, []).append(right_ns)

        joined_rows: list[dict[str, dict[str, Any]]] = []
        right_null_values = [None for _ in right_headers]
        right_null_ns = self._build_source_row(right_source, right_headers, right_null_values)

        if join_type.upper() == "RIGHT":
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
                        combined_row: dict[str, dict[str, Any]] = {}
                        combined_row.update(left_ns)
                        combined_row.update(right_ns)
                        joined_rows.append(combined_row)
                else:
                    combined_row = {}
                    combined_row.update(left_null_ns)
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
                elif join_type.upper() == "LEFT":
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

            for clause in join_spec["on"]["clauses"]:
                for side in ("left", "right"):
                    col = clause[side]
                    _validate_column_ref(str(col["source"]), str(col["name"]), "ON")

            known_sources.update({right_ref, right_table_name})
            join_inputs.append((join_spec, right_data))

        # Validate SELECT columns
        for column in parsed["columns"]:
            if isinstance(column, dict) and column.get("type") == "column":
                _validate_column_ref(str(column["source"]), str(column["name"]), "SELECT")

        # Validate WHERE columns
        where_raw = parsed.get("where")
        if where_raw:
            self._validate_join_where_refs(where_raw, _validate_column_ref)

        # Validate ORDER BY column
        order_by = self._normalize_order_by(parsed.get("order_by"))
        if order_by:
            for item in order_by:
                col_ref = str(item["column"])
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
            left_rows, left_headers_map = self._join_two_sources(
                left_rows=left_rows,
                left_headers_map=left_headers_map,
                right_data=right_data,
                right_source=right_source,
                join_type=str(join_spec["type"]),
                on_clauses=join_spec["on"]["clauses"],
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

        order_by = self._normalize_order_by(parsed.get("order_by"))
        if order_by:
            joined_rows_flat = self._apply_order_by(
                joined_rows_flat,
                order_by,
                value_getter=lambda r, col: r.get(col),
            )

        offset = parsed.get("offset") or 0
        limit = parsed.get("limit")
        if offset:
            joined_rows_flat = joined_rows_flat[offset:]
        if limit is not None:
            joined_rows_flat = joined_rows_flat[:limit]

        selected_columns: list[str] = []
        for column in parsed["columns"]:
            source = str(column["source"])
            name = str(column["name"])
            selected_columns.append(f"{source}.{name}")

        rows_out = [
            tuple(row.get(column_name) for column_name in selected_columns)
            for row in joined_rows_flat
        ]
        description: Description = [
            (column_name, None, None, None, None, None, None)
            for column_name in selected_columns
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
        selected_columns = headers if columns == ["*"] else list(columns)
        missing = [col for col in selected_columns if col not in headers]
        if missing:
            raise ValueError(
                f"Unknown column(s): {', '.join(missing)}. Available columns: {headers}"
            )

        order_by = self._normalize_order_by(parsed.get("order_by"))
        if order_by:
            rows = self._apply_order_by(
                rows,
                order_by,
                value_getter=lambda r, col: r.get(col),
                available_columns=set(headers),
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
        required_aggregates: dict[str, tuple[str, str]] = {}
        for column in columns:
            if self._is_aggregate_column(column):
                arg = str(column.get("arg"))
                if arg != "*" and arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                label = self._aggregate_label(column)
                required_aggregates[label] = (str(column.get("func")), arg)
                output_columns.append(label)
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

        if having is not None:
            if "conditions" in having:
                refs = [str(condition["column"]) for condition in having["conditions"]]
            else:
                refs = [str(having["column"])]
            for ref in refs:
                aggregate_spec = self._aggregate_spec_from_label(ref)
                if aggregate_spec is None:
                    continue
                func, arg = aggregate_spec
                if arg != "*" and arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                required_aggregates[ref] = (func, arg)

        order_by_clause_raw = parsed.get("order_by")
        if isinstance(order_by_clause_raw, list):
            order_by_clause = order_by_clause_raw
        elif isinstance(order_by_clause_raw, dict):
            order_by_clause = [order_by_clause_raw]
        else:
            order_by_clause = None
        if order_by_clause is not None:
            for item in order_by_clause:
                ref = str(item["column"])
                aggregate_spec = self._aggregate_spec_from_label(ref)
                if aggregate_spec is None:
                    continue
                func, arg = aggregate_spec
                if arg != "*" and arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                required_aggregates[ref] = (func, arg)

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
                for label, (func, arg) in required_aggregates.items():
                    context_row[label] = self._compute_aggregate(func, arg, group_values)
                if having and not self._matches_where(context_row, having):
                    continue
                grouped_rows.append(context_row)
        else:
            context_row_single: dict[str, Any] = {}
            for label, (func, arg) in required_aggregates.items():
                context_row_single[label] = self._compute_aggregate(func, arg, rows)
            if not having or self._matches_where(context_row_single, having):
                grouped_rows.append(context_row_single)

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

        order_by = self._normalize_order_by(parsed.get("order_by"))
        if order_by:
            available_order_columns = set(output_columns)
            if group_by:
                available_order_columns.update(group_by)
            available_order_columns.update(required_aggregates.keys())
            grouped_rows = self._apply_order_by(
                grouped_rows,
                order_by,
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

    def _aggregate_spec_from_label(self, expression: str) -> tuple[str, str] | None:
        match = re.fullmatch(r"(?i)(COUNT|SUM|AVG|MIN|MAX)\(([^\)]+)\)", expression.strip())
        if not match:
            return None
        func = match.group(1).upper()
        arg = match.group(2).strip()
        if not arg:
            return None
        if arg != "*" and not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", arg):
            raise ValueError(
                f"Unsupported aggregate expression: {func}({arg}). "
                "Only bare column names and * are supported"
            )
        return func, arg

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
        if operator == "NOT IN":
            if row_value is None:
                return False
            for candidate in value:
                left, right = self._coerce_for_compare(row_value, candidate)
                if left == right:
                    return False
            return True
        if operator == "BETWEEN":
            if row_value is None:
                return False
            low, high = value
            if low is None or high is None:
                return False
            left_low, right_low = self._coerce_for_compare(row_value, low)
            left_high, right_high = self._coerce_for_compare(row_value, high)
            return bool(left_low >= right_low and left_high <= right_high)
        if operator == "NOT BETWEEN":
            if row_value is None:
                return False
            low, high = value
            if low is None or high is None:
                return False
            left_low, right_low = self._coerce_for_compare(row_value, low)
            left_high, right_high = self._coerce_for_compare(row_value, high)
            return not bool(left_low >= right_low and left_high <= right_high)
        if operator == "LIKE":
            if row_value is None:
                return False
            if not isinstance(value, str):
                raise NotImplementedError("Unsupported LIKE pattern type")
            regex = "^" + re.escape(value).replace(r"%", ".*").replace(r"_", ".") + "$"
            return bool(re.match(regex, str(row_value)))
        if operator == "NOT LIKE":
            if row_value is None:
                return False
            if not isinstance(value, str):
                raise NotImplementedError("Unsupported LIKE pattern type")
            regex = "^" + re.escape(value).replace(r"%", ".*").replace(r"_", ".") + "$"
            return not bool(re.match(regex, str(row_value)))

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
