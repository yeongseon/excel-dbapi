import copy
from datetime import date, datetime, time
import importlib
import logging
import re
import warnings
from typing import Any, Callable, Iterator, Protocol, cast

from .engines.base import TableData, WorkbookBackend
from .engines.result import Description, ExecutionResult
from .exceptions import ProgrammingError
from .parser import _parse_column_expression, parse_sql
from .reflection import METADATA_SHEET
from .sanitize import sanitize_cell_value, sanitize_row

_READONLY_ACTIONS = frozenset({"INSERT", "UPDATE", "DELETE", "CREATE", "DROP", "ALTER"})

_logger = logging.getLogger(__name__)


class _SupportsOrder(Protocol):
    def __lt__(self, other: Any, /) -> bool: ...


def _build_like_regex(pattern: str, escape_char: str | None) -> str:
    parts: list[str] = ["^"]
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if escape_char is not None and char == escape_char:
            index += 1
            if index >= len(pattern):
                raise ValueError("Invalid LIKE pattern: trailing ESCAPE character")
            parts.append(re.escape(pattern[index]))
        elif char == "%":
            parts.append(".*")
        elif char == "_":
            parts.append(".")
        else:
            parts.append(re.escape(char))
        index += 1

    parts.append("$")
    return "".join(parts)


ScalarFunctionHandler = Callable[[list[Any]], Any]
ScalarFunctionSpec = tuple[int, int | None, ScalarFunctionHandler]


def _coalesce(args: list[Any]) -> Any:
    for value in args:
        if value is not None:
            return value
    return None


def _nullif(args: list[Any]) -> Any:
    return None if args[0] == args[1] else args[0]


def _upper(args: list[Any]) -> Any:
    value = args[0]
    return None if value is None else str(value).upper()


def _lower(args: list[Any]) -> Any:
    value = args[0]
    return None if value is None else str(value).lower()


def _trim(args: list[Any]) -> Any:
    value = args[0]
    return None if value is None else str(value).strip()


def _length(args: list[Any]) -> Any:
    value = args[0]
    return None if value is None else len(str(value))


def _to_int_like(value: Any) -> int:
    if isinstance(value, bool):
        raise ValueError("expected numeric value")
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError("expected numeric value")
        return int(float(text))
    raise ValueError("expected numeric value")


def _substr(args: list[Any]) -> Any:
    text_value = args[0]
    start_value = args[1]
    if text_value is None or start_value is None:
        return None

    text = str(text_value)
    start = _to_int_like(start_value)
    if start > 0:
        start_index = start - 1
    elif start < 0:
        start_index = len(text) + start
    else:
        start_index = 0

    if len(args) < 3:
        return text[start_index:]

    length_value = args[2]
    if length_value is None:
        return None
    length = _to_int_like(length_value)
    if length <= 0:
        return ""
    return text[start_index : start_index + length]


def _concat(args: list[Any]) -> str:
    return "".join(str(value) for value in args if value is not None)


def _abs(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError("expected numeric value")
    if isinstance(value, (int, float)):
        return abs(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError("expected numeric value")
        return abs(float(text))
    raise ValueError("expected numeric value")


def _round(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError("expected numeric value")
    if isinstance(value, (int, float)):
        numeric = float(value)
    elif isinstance(value, str):
        text = value.strip()
        if not text:
            raise ValueError("expected numeric value")
        numeric = float(text)
    else:
        raise ValueError("expected numeric value")

    if len(args) < 2 or args[1] is None:
        return round(numeric)
    precision = _to_int_like(args[1])
    return round(numeric, precision)


def _replace(args: list[Any]) -> Any:
    source = args[0]
    if source is None:
        return None
    old = args[1]
    if old is None:
        return str(source)
    new = args[2]
    return str(source).replace(str(old), "" if new is None else str(new))


def _date_value(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value.replace(tzinfo=None) if value.tzinfo is not None else value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if isinstance(value, str):
        normalized = value.strip()
        if not normalized:
            raise ValueError("expected date value")
        if normalized.endswith("Z"):
            normalized = normalized[:-1] + "+00:00"
        try:
            parsed_datetime = datetime.fromisoformat(normalized)
        except ValueError:
            parsed_datetime = None
        if parsed_datetime is not None:
            return (
                parsed_datetime.replace(tzinfo=None)
                if parsed_datetime.tzinfo is not None
                else parsed_datetime
            )
        try:
            parsed_date = date.fromisoformat(value.strip())
        except ValueError as exc:
            raise ValueError("expected date value") from exc
        return datetime.combine(parsed_date, time.min)
    raise ValueError("expected date value")


def _year(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    return _date_value(value).year


def _month(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    return _date_value(value).month


def _day(args: list[Any]) -> Any:
    value = args[0]
    if value is None:
        return None
    return _date_value(value).day


_SCALAR_FUNCTIONS: dict[str, ScalarFunctionSpec] = {
    "COALESCE": (1, None, _coalesce),
    "NULLIF": (2, 2, _nullif),
    "UPPER": (1, 1, _upper),
    "LOWER": (1, 1, _lower),
    "TRIM": (1, 1, _trim),
    "LENGTH": (1, 1, _length),
    "SUBSTR": (2, 3, _substr),
    "SUBSTRING": (2, 3, _substr),
    "ABS": (1, 1, _abs),
    "ROUND": (1, 2, _round),
    "REPLACE": (3, 3, _replace),
    "CONCAT": (1, None, _concat),
    "YEAR": (1, 1, _year),
    "MONTH": (1, 1, _month),
    "DAY": (1, 1, _day),
}


def _tv_and(a: bool | None, b: bool | None) -> bool | None:
    """SQL three-valued AND."""
    if a is False or b is False:
        return False
    if a is None or b is None:
        return None
    return True


def _tv_or(a: bool | None, b: bool | None) -> bool | None:
    """SQL three-valued OR."""
    if a is True or b is True:
        return True
    if a is None or b is None:
        return None
    return False


class SharedExecutor:
    def __init__(
        self,
        backend: WorkbookBackend,
        *,
        sanitize_formulas: bool = True,
        connection: Any | None = None,
    ):
        self.backend = backend
        self.sanitize_formulas = sanitize_formulas
        self._connection = connection
        self._subquery_cache: dict[int, Any] = {}
        self._outer_row_stack: list[dict[str, Any]] = []
        self._cte_tables: dict[str, TableData] = {}

    def _write_metadata_for_headers(
        self,
        table_name: str,
        headers: list[str],
        type_by_column: dict[str, str] | None = None,
    ) -> None:
        if self._connection is None:
            return
        reflection_module = importlib.import_module("excel_dbapi.reflection")

        existing_type_by_column: dict[str, str] = {}
        try:
            existing_metadata = reflection_module.read_table_metadata(
                self._connection,
                table_name,
            )
        except Exception as exc:
            if getattr(self.backend, "supports_transactions", True):
                raise
            _logger.warning(
                "Metadata read skipped before non-transactional metadata write (%s): %s",
                table_name,
                exc,
            )
            existing_metadata = None
            normalized_type_by_column = (
                {key.casefold(): value for key, value in type_by_column.items()}
                if type_by_column is not None
                else {}
            )
            has_complete_type_map = all(
                header.casefold() in normalized_type_by_column for header in headers
            )
            if not has_complete_type_map:
                warnings.warn(
                    (
                        f"Could not read metadata for '{table_name}'; "
                        "skipping metadata update to avoid data loss"
                    ),
                    UserWarning,
                    stacklevel=2,
                )
                return

        if existing_metadata is not None:
            existing_type_by_column = {
                str(entry["name"]).casefold(): str(entry.get("type_name", "TEXT"))
                for entry in existing_metadata
            }

        normalized_type_by_column = (
            {key.casefold(): value for key, value in type_by_column.items()}
            if type_by_column is not None
            else {}
        )

        columns = [
            {
                "name": header,
                "type_name": normalized_type_by_column.get(
                    header.casefold(),
                    existing_type_by_column.get(header.casefold(), "TEXT"),
                ),
                "nullable": True,
                "primary_key": False,
            }
            for header in headers
        ]
        self._sync_metadata_write(
            lambda: reflection_module.write_table_metadata(
                self._connection,
                table_name,
                columns,
            ),
            context=f"write metadata for table '{table_name}'",
        )

    def _sync_metadata_write(
        self, operation: Callable[[], None], *, context: str
    ) -> None:
        try:
            operation()
        except Exception as exc:
            if getattr(self.backend, "supports_transactions", True):
                raise
            _logger.warning(
                "Metadata sync skipped after non-transactional workbook mutation (%s): %s",
                context,
                exc,
            )

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

    def execute(
        self,
        parsed: dict[str, Any],
        *,
        _reset_subquery_cache: bool = True,
    ) -> ExecutionResult:
        if _reset_subquery_cache:
            self._subquery_cache.clear()

        ctes = parsed.get("ctes")
        if isinstance(ctes, list) and ctes:
            previous_ctes = dict(self._cte_tables)
            try:
                for cte in ctes:
                    cte_name = cte.get("name")
                    cte_query = cte.get("query")
                    if not isinstance(cte_name, str) or not isinstance(cte_query, dict):
                        raise ValueError("Invalid CTE definition")
                    cte_result = self.execute(cte_query, _reset_subquery_cache=False)
                    self._cte_tables[cte_name] = TableData(
                        headers=[str(col[0]) for col in cte_result.description],
                        rows=[list(row) for row in cte_result.rows],
                    )

                main_query = dict(parsed)
                main_query.pop("ctes", None)
                return self.execute(main_query, _reset_subquery_cache=False)
            finally:
                self._cte_tables = previous_ctes

        action = parsed["action"]
        self._ensure_writable(action)

        if action == "COMPOUND":
            return self._execute_compound(parsed)

        table = parsed["table"]
        resolved_table = self._resolve_sheet_name(table)

        if action == "SELECT":
            selected_table, selected_data = self._resolve_table_data(table)
            if selected_table is None or selected_data is None:
                available = self._available_table_names()
                msg = f"Sheet '{table}' not found in Excel."
                if available:
                    msg += f" Available sheets: {available}"
                raise ValueError(msg)
            if parsed.get("joins") is not None:
                return self._execute_join_select(
                    action, parsed, selected_table, selected_data
                )
            if not selected_data.headers:
                if parsed.get("columns") != ["*"]:
                    raise ValueError(
                        f"No columns defined in sheet '{selected_table}' — cannot resolve column references"
                    )
                return ExecutionResult(
                    action=action, rows=[], description=[], rowcount=0, lastrowid=None
                )
            headers = list(selected_data.headers)
            source_refs: set[str] = set()
            from_entry = parsed.get("from")
            if isinstance(from_entry, dict):
                table_name = from_entry.get("table")
                if isinstance(table_name, str):
                    source_refs.add(table_name)
                ref_name = from_entry.get("ref")
                if isinstance(ref_name, str):
                    source_refs.add(ref_name)

            rows = [
                self._build_scoped_row(
                    self._row_from_values(headers, list(row_values)),
                    headers=headers,
                    source_refs=source_refs,
                )
                for row_values in selected_data.rows
            ]
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
                raise ValueError(
                    f"No columns defined in sheet '{resolved_table}' — cannot resolve column references"
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
                scoped_row = self._build_scoped_row(
                    row_map,
                    headers=headers,
                    source_refs={table},
                )
                if where and not self._matches_where(scoped_row, where):
                    continue
                for update in updates:
                    col_index = headers.index(update["column"])
                    raw_value = update["value"]
                    should_eval_str = (
                        isinstance(raw_value, str) and raw_value in scoped_row
                    )
                    if should_eval_str or (
                        isinstance(raw_value, dict)
                        and raw_value.get("type")
                        in {
                            "alias",
                            "case",
                            "binary_op",
                            "unary_op",
                            "literal",
                            "column",
                            "function",
                            "cast",
                        }
                    ):
                        evaluated = self._eval_expression(
                            raw_value,
                            scoped_row,
                            lambda c: scoped_row.get(c),
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
                where = parsed.get("where")
                if where and self._collect_where_column_refs(where):
                    raise ValueError(
                        f"No columns defined in sheet '{resolved_table}' — cannot resolve column references"
                    )
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
                scoped_row = self._build_scoped_row(
                    row_map,
                    headers=headers,
                    source_refs={table},
                )
                if where and not self._matches_where(scoped_row, where):
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
                subquery_result = self.execute(
                    values["query"],
                    _reset_subquery_cache=False,
                )
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
                        raise ValueError(
                            "INSERT values count does not match header count"
                        )
                    else:
                        raise ValueError(
                            "INSERT values count does not match column count"
                        )
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
                        raise ValueError(
                            f"ON CONFLICT column '{target_col}' not found in headers"
                        )

                target_indices = [
                    headers.index(target_col) for target_col in target_cols
                ]
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
                            left, right = self._coerce_for_compare(
                                existing_value, incoming_value
                            )
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
                        should_eval_str = isinstance(raw_value, str) and (
                            raw_value in row_map
                            or (
                                raw_value.startswith("excluded.")
                                and raw_value.split(".", 1)[1] in excluded_map
                            )
                        )
                        if should_eval_str or (
                            isinstance(raw_value, dict)
                            and raw_value.get("type")
                            in {
                                "alias",
                                "case",
                                "binary_op",
                                "unary_op",
                                "literal",
                                "column",
                                "function",
                                "cast",
                            }
                        ):
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
                            conflict_row.extend(
                                [None] * (col_index - len(conflict_row) + 1)
                            )
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
            if table.casefold() == METADATA_SHEET.casefold():
                raise ProgrammingError(
                    "Cannot perform DDL on reserved metadata table '__excel_meta__'"
                )
            if resolved_table is not None:
                raise ValueError(f"Sheet '{table}' already exists")
            columns = parsed["columns"]
            seen: set[str] = set()
            for col in columns:
                lower = col.lower()
                if lower in seen:
                    raise ValueError(f"Duplicate column name '{col}' in CREATE TABLE")
                seen.add(lower)
            self.backend.create_sheet(table, columns)
            type_by_column = {
                str(definition["name"]): str(definition.get("type_name", "TEXT"))
                for definition in parsed.get("column_definitions", [])
            }
            self._write_metadata_for_headers(table, columns, type_by_column)
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        if action == "DROP":
            if table.casefold() == METADATA_SHEET.casefold():
                raise ProgrammingError(
                    "Cannot perform DDL on reserved metadata table '__excel_meta__'"
                )
            if resolved_table is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            reflection_module = importlib.import_module("excel_dbapi.reflection")

            user_sheets = [
                sheet_name
                for sheet_name in self.backend.list_sheets()
                if sheet_name != METADATA_SHEET
            ]
            if len(user_sheets) <= 1:
                raise ValueError("Cannot drop the only remaining sheet")
            self.backend.drop_sheet(resolved_table)
            if self._connection is not None:
                self._sync_metadata_write(
                    lambda: reflection_module.remove_table_metadata(
                        self._connection,
                        resolved_table,
                    ),
                    context=f"remove metadata for table '{resolved_table}'",
                )
            return ExecutionResult(
                action=action,
                rows=[],
                description=[],
                rowcount=0,
                lastrowid=None,
            )

        if action == "ALTER":
            if table.casefold() == METADATA_SHEET.casefold():
                raise ProgrammingError(
                    "Cannot perform DDL on reserved metadata table '__excel_meta__'"
                )
            if resolved_table is None:
                raise ValueError(f"Sheet '{table}' not found in Excel")
            operation = parsed.get("operation")
            data = self.backend.read_sheet(resolved_table)

            if operation == "ADD_COLUMN":
                col = parsed["column"]
                if col in data.headers or col.lower() in {
                    h.lower() for h in data.headers
                }:
                    raise ValueError(f"Column '{col}' already exists in '{table}'")
                data.headers.append(col)
                for row in data.rows:
                    row.append(None)
                self.backend.write_sheet(resolved_table, data)
                self._write_metadata_for_headers(
                    resolved_table,
                    list(data.headers),
                    {col: str(parsed.get("type_name", "TEXT"))},
                )
            elif operation == "DROP_COLUMN":
                col = parsed["column"]
                target = col.casefold()
                idx = next(
                    (
                        i
                        for i, header in enumerate(data.headers)
                        if header.casefold() == target
                    ),
                    -1,
                )
                if idx == -1:
                    raise ValueError(f"Column '{col}' not found in '{table}'")
                if len(data.headers) == 1:
                    raise ValueError(
                        f"Cannot drop the only column '{col}' from '{table}'"
                    )
                data.headers.pop(idx)
                for row in data.rows:
                    if idx < len(row):
                        row.pop(idx)
                self.backend.write_sheet(resolved_table, data)
                self._write_metadata_for_headers(resolved_table, list(data.headers))
            elif operation == "RENAME_COLUMN":
                old_col = parsed["old_column"]
                new_col = parsed["new_column"]
                target = old_col.casefold()
                idx = next(
                    (
                        i
                        for i, header in enumerate(data.headers)
                        if header.casefold() == target
                    ),
                    -1,
                )
                if idx == -1:
                    raise ValueError(f"Column '{old_col}' not found in '{table}'")
                matched_old_col = data.headers[idx]
                if new_col in data.headers or new_col.casefold() in {
                    h.casefold() for h in data.headers if h != matched_old_col
                }:
                    raise ValueError(f"Column '{new_col}' already exists in '{table}'")
                data.headers[idx] = new_col
                self.backend.write_sheet(resolved_table, data)
                self._write_metadata_for_headers(resolved_table, list(data.headers))
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
        return tuple(
            tuple(value) if isinstance(value, list) else value for value in row
        )

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

    def _dedupe_projected_rows(
        self,
        rows: list[dict[str, Any]],
        projected_columns: list[str],
    ) -> list[dict[str, Any]]:
        seen: set[tuple[Any, ...]] = set()
        deduped: list[dict[str, Any]] = []
        for row in rows:
            key = self._normalize_row_key(
                tuple(row.get(column_name) for column_name in projected_columns)
            )
            if key in seen:
                continue
            seen.add(key)
            deduped.append(row)
        return deduped

    @staticmethod
    def _validate_distinct_order_by_columns(
        order_by: list[dict[str, Any]] | None,
        selected_columns: list[str],
    ) -> None:
        if not order_by:
            return
        selected = set(selected_columns)
        for item in order_by:
            column_name = str(item["column"])
            if column_name not in selected:
                raise ValueError(
                    "ORDER BY columns must appear in SELECT list when using DISTINCT"
                )

    @staticmethod
    def _normalize_single_source_aggregate_arg(arg: str, headers: list[str]) -> str:
        if arg == "*" or "." not in arg:
            return arg
        if arg in headers:
            return arg
        if any("." in header for header in headers):
            return arg
        _, bare_column = arg.split(".", 1)
        if bare_column in headers:
            return bare_column
        return arg

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
                return SharedExecutor._expression_label(column)
            if column.get("type") == "column":
                return f"{column['source']}.{column['name']}"
            if column.get("type") in {
                "binary_op",
                "unary_op",
                "literal",
                "function",
                "cast",
                "subquery",
                "window_function",
            }:
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
                return SharedExecutor._expression_to_sql(inner)
            if inner.get("type") == "column":
                return f"{inner['source']}.{inner['name']}"
            if inner.get("type") in {
                "binary_op",
                "unary_op",
                "literal",
                "case",
                "function",
                "cast",
                "subquery",
                "window_function",
            }:
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
                func = str(expression.get("func", "")).upper()
                arg = str(expression.get("arg", "")).strip()
                aggregate_sql = (
                    f"{func}(DISTINCT {arg})"
                    if expression.get("distinct")
                    else f"{func}({arg})"
                )
                filter_clause = expression.get("filter")
                if isinstance(filter_clause, dict):
                    filter_sql = SharedExecutor._where_to_sql(filter_clause)
                    aggregate_sql = f"{aggregate_sql} FILTER (WHERE {filter_sql})"
                return aggregate_sql
            if expression_type == "window_function":
                func = str(expression.get("func", "")).upper()
                args = expression.get("args")
                args_list = args if isinstance(args, list) else []
                args_sql_parts = [
                    SharedExecutor._expression_to_sql(argument)
                    if isinstance(argument, dict)
                    else str(argument)
                    for argument in args_list
                ]
                args_sql = ", ".join(args_sql_parts)
                if expression.get("distinct") and args_sql:
                    args_sql = f"DISTINCT {args_sql}"

                function_sql = f"{func}({args_sql})"
                filter_clause = expression.get("filter")
                if isinstance(filter_clause, dict):
                    filter_sql = SharedExecutor._where_to_sql(filter_clause)
                    function_sql = f"{function_sql} FILTER (WHERE {filter_sql})"

                spec_parts: list[str] = []
                partition_by = expression.get("partition_by")
                if isinstance(partition_by, list) and partition_by:
                    partition_sql = ", ".join(
                        SharedExecutor._expression_to_sql(partition_expression)
                        for partition_expression in partition_by
                    )
                    spec_parts.append(f"PARTITION BY {partition_sql}")

                order_by = expression.get("order_by")
                if isinstance(order_by, list) and order_by:
                    order_parts: list[str] = []
                    for order_item in order_by:
                        if not isinstance(order_item, dict):
                            continue
                        order_expression = order_item.get("__expression__")
                        if order_expression is not None:
                            order_sql = SharedExecutor._expression_to_sql(
                                order_expression
                            )
                        else:
                            order_sql = str(order_item.get("column", ""))
                            if order_sql.startswith("__expr__:"):
                                order_sql = order_sql[len("__expr__:") :]
                        direction = str(order_item.get("direction", "ASC")).upper()
                        order_parts.append(f"{order_sql} {direction}")
                    if order_parts:
                        spec_parts.append("ORDER BY " + ", ".join(order_parts))

                if spec_parts:
                    return f"{function_sql} OVER ({' '.join(spec_parts)})"
                return f"{function_sql} OVER ()"
            if expression_type == "literal":
                return SharedExecutor._literal_to_sql(expression.get("value"))
            if expression_type == "unary_op":
                operand_sql = SharedExecutor._expression_to_sql(
                    expression.get("operand")
                )
                return f"-{operand_sql}"
            if expression_type == "binary_op":
                left_sql = SharedExecutor._expression_to_sql(expression.get("left"))
                right_sql = SharedExecutor._expression_to_sql(expression.get("right"))
                return f"({left_sql} {expression['op']} {right_sql})"
            if expression_type == "function":
                args = expression.get("args")
                args_list = args if isinstance(args, list) else []
                args_sql = ", ".join(
                    SharedExecutor._expression_to_sql(arg) for arg in args_list
                )
                return f"{expression['name']}({args_sql})"
            if expression_type == "cast":
                value_sql = SharedExecutor._expression_to_sql(expression.get("value"))
                target_type = str(expression.get("target_type", ""))
                return f"CAST({value_sql} AS {target_type})"
            if expression_type == "subquery":
                return "(SUBQUERY)"
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
                        match_sql = SharedExecutor._expression_to_sql(
                            when_branch.get("match")
                        )
                        parts.append(f"WHEN {match_sql} THEN")
                    parts.append(
                        SharedExecutor._expression_to_sql(when_branch["result"])
                    )
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
            if operand.get("type") == "exists":
                return "EXISTS (SUBQUERY)"
            return SharedExecutor._expression_to_sql(operand)
        if isinstance(operand, str):
            if is_column:
                return operand
            return SharedExecutor._literal_to_sql(operand)
        return SharedExecutor._literal_to_sql(operand)

    @staticmethod
    def _where_to_sql(where: dict[str, Any]) -> str:
        node_type = where.get("type")
        if node_type == "exists":
            return "EXISTS (SUBQUERY)"
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
                values_sql = SharedExecutor._where_operand_to_sql(
                    value, is_column=False
                )
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
        if expression_type in {
            "binary_op",
            "unary_op",
            "literal",
            "case",
            "function",
            "cast",
            "subquery",
            "window_function",
        }:
            return True
        if expression_type == "alias":
            return SharedExecutor._contains_arithmetic_expression(
                inner.get("expression")
            )
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
            return SharedExecutor._collect_expression_column_refs(
                expression.get("expression")
            )
        if expression_type == "column":
            refs.add(f"{expression['source']}.{expression['name']}")
            return refs
        if expression_type == "aggregate":
            arg = str(expression.get("arg", ""))
            if arg and arg != "*":
                refs.add(arg)
            filter_clause = expression.get("filter")
            if isinstance(filter_clause, dict):
                refs.update(SharedExecutor._collect_where_column_refs(filter_clause))
            return refs
        if expression_type == "window_function":
            args = expression.get("args")
            if isinstance(args, list):
                for argument in args:
                    if isinstance(argument, str):
                        if argument and argument != "*":
                            refs.add(argument)
                    else:
                        refs.update(
                            SharedExecutor._collect_expression_column_refs(argument)
                        )

            partition_by = expression.get("partition_by")
            if isinstance(partition_by, list):
                for partition_expression in partition_by:
                    refs.update(
                        SharedExecutor._collect_expression_column_refs(
                            partition_expression
                        )
                    )

            order_by = expression.get("order_by")
            if isinstance(order_by, list):
                for order_item in order_by:
                    if not isinstance(order_item, dict):
                        continue
                    order_expression = order_item.get("__expression__")
                    if order_expression is not None:
                        refs.update(
                            SharedExecutor._collect_expression_column_refs(
                                order_expression
                            )
                        )
                        continue
                    order_column = order_item.get("column")
                    if isinstance(order_column, str) and order_column:
                        refs.add(order_column)

            filter_clause = expression.get("filter")
            if isinstance(filter_clause, dict):
                refs.update(SharedExecutor._collect_where_column_refs(filter_clause))
            return refs
        if expression_type == "binary_op":
            refs.update(
                SharedExecutor._collect_expression_column_refs(expression.get("left"))
            )
            refs.update(
                SharedExecutor._collect_expression_column_refs(expression.get("right"))
            )
            return refs
        if expression_type == "unary_op":
            refs.update(
                SharedExecutor._collect_expression_column_refs(
                    expression.get("operand")
                )
            )
            return refs
        if expression_type == "function":
            args = expression.get("args")
            if isinstance(args, list):
                for arg in args:
                    refs.update(SharedExecutor._collect_expression_column_refs(arg))
            return refs
        if expression_type == "cast":
            refs.update(
                SharedExecutor._collect_expression_column_refs(expression.get("value"))
            )
            return refs
        if expression_type == "subquery":
            return refs
        if expression_type == "case":
            if expression.get("value") is not None:
                refs.update(
                    SharedExecutor._collect_expression_column_refs(expression["value"])
                )
            for when_branch in expression.get("whens", []):
                # Collect from condition (searched mode) or match (simple mode)
                condition = when_branch.get("condition")
                if condition is not None:
                    refs.update(SharedExecutor._collect_where_column_refs(condition))
                match_expr = when_branch.get("match")
                if match_expr is not None:
                    refs.update(
                        SharedExecutor._collect_expression_column_refs(match_expr)
                    )
                refs.update(
                    SharedExecutor._collect_expression_column_refs(
                        when_branch["result"]
                    )
                )
            if expression.get("else") is not None:
                refs.update(
                    SharedExecutor._collect_expression_column_refs(expression["else"])
                )
            return refs
        return refs

    @staticmethod
    def _collect_where_column_refs(where: dict[str, Any]) -> set[str]:
        """Collect column references from a WHERE condition tree."""
        refs: set[str] = set()
        node_type = where.get("type")
        if node_type == "exists":
            return refs
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
        if isinstance(value, dict) and value.get("type") not in {"subquery", "exists"}:
            refs.update(SharedExecutor._collect_expression_column_refs(value))
        elif isinstance(value, (list, tuple)):
            for candidate in value:
                if isinstance(candidate, dict):
                    refs.update(
                        SharedExecutor._collect_expression_column_refs(candidate)
                    )
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
            expression_ast = item.get("__expression__")
            if expression_ast is not None:
                parsed_expressions[col_ref] = expression_ast
                continue

            expression_sql = col_ref[len("__expr__:") :]
            parsed_expressions[col_ref] = _parse_column_expression(
                expression_sql,
                allow_wildcard=False,
                allow_aggregates=False,
                allow_subqueries=True,
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

    @staticmethod
    def _collect_window_expressions(
        expression: Any,
        collected: dict[str, dict[str, Any]],
    ) -> None:
        if not isinstance(expression, dict):
            return

        expression_type = expression.get("type")
        if expression_type == "alias":
            SharedExecutor._collect_window_expressions(
                expression.get("expression"), collected
            )
            return

        if expression_type == "window_function":
            collected[SharedExecutor._source_key(expression)] = expression
            return

        if expression_type == "unary_op":
            SharedExecutor._collect_window_expressions(
                expression.get("operand"), collected
            )
            return

        if expression_type == "binary_op":
            SharedExecutor._collect_window_expressions(
                expression.get("left"), collected
            )
            SharedExecutor._collect_window_expressions(
                expression.get("right"), collected
            )
            return

        if expression_type == "function":
            args = expression.get("args")
            if isinstance(args, list):
                for argument in args:
                    SharedExecutor._collect_window_expressions(argument, collected)
            return

        if expression_type == "cast":
            SharedExecutor._collect_window_expressions(
                expression.get("value"), collected
            )
            return

        if expression_type == "case":
            SharedExecutor._collect_window_expressions(
                expression.get("value"), collected
            )
            for when_branch in expression.get("whens", []):
                if not isinstance(when_branch, dict):
                    continue
                SharedExecutor._collect_window_expressions(
                    when_branch.get("match"), collected
                )
                SharedExecutor._collect_window_expressions(
                    when_branch.get("result"), collected
                )
            SharedExecutor._collect_window_expressions(
                expression.get("else"), collected
            )

    def _apply_window_functions(
        self,
        rows: list[dict[str, Any]],
        columns: list[Any],
        order_by: list[dict[str, Any]] | None,
    ) -> set[str]:
        window_expressions: dict[str, dict[str, Any]] = {}

        if columns != ["*"]:
            for column in columns:
                inner = self._unwrap_alias(column)
                self._collect_window_expressions(inner, window_expressions)

        if order_by:
            for item in order_by:
                expression = item.get("__expression__")
                if expression is not None:
                    self._collect_window_expressions(expression, window_expressions)

        if not window_expressions:
            return set()

        for target_column, expression in window_expressions.items():
            self._evaluate_window_expression(
                rows, expression, target_column=target_column
            )

        return set(window_expressions.keys())

    def _evaluate_window_expression(
        self,
        rows: list[dict[str, Any]],
        expression: dict[str, Any],
        *,
        target_column: str,
    ) -> None:
        function_name = str(expression.get("func", "")).upper()
        partition_by = expression.get("partition_by")
        partition_expressions = partition_by if isinstance(partition_by, list) else []
        order_by = self._normalize_order_by(expression.get("order_by")) or []

        partitions: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
        for row in rows:
            partition_values = [
                self._eval_expression(
                    partition_expression,
                    row,
                    lambda col_name: row.get(col_name),
                )
                for partition_expression in partition_expressions
            ]
            partition_key = self._normalize_row_key(tuple(partition_values))
            partitions.setdefault(partition_key, []).append(row)

        distinct = bool(expression.get("distinct"))
        filter_clause = expression.get("filter")
        filter_condition = filter_clause if isinstance(filter_clause, dict) else None
        args = expression.get("args")
        args_list = args if isinstance(args, list) else []

        for partition_rows in partitions.values():
            ordered_rows = partition_rows
            if order_by:
                self._materialize_order_expression_columns(partition_rows, order_by)
                ordered_rows = self._apply_order_by(
                    partition_rows,
                    order_by,
                    value_getter=lambda candidate, column_name: candidate.get(
                        column_name
                    ),
                )

            if function_name == "ROW_NUMBER":
                for position, row in enumerate(ordered_rows, start=1):
                    row[target_column] = position
                continue

            if function_name in {"RANK", "DENSE_RANK"}:
                if not order_by:
                    for row in ordered_rows:
                        row[target_column] = 1
                    continue

                previous_key: tuple[Any, ...] | None = None
                rank = 1
                dense_rank = 1
                for position, row in enumerate(ordered_rows, start=1):
                    current_key = tuple(
                        self._sort_key(row.get(str(item["column"])))
                        for item in order_by
                    )
                    if previous_key is None:
                        rank = 1
                        dense_rank = 1
                    elif current_key != previous_key:
                        rank = position
                        dense_rank += 1

                    row[target_column] = rank if function_name == "RANK" else dense_rank
                    previous_key = current_key
                continue

            if function_name in {"COUNT", "SUM", "AVG", "MIN", "MAX"}:
                if not args_list:
                    raise ProgrammingError(
                        f"Window function {function_name} requires an argument"
                    )
                arg = str(args_list[0])
                if not order_by:
                    partition_value = self._compute_aggregate(
                        function_name,
                        arg,
                        ordered_rows,
                        distinct=distinct,
                        filter_condition=filter_condition,
                    )
                    for row in ordered_rows:
                        row[target_column] = partition_value
                    continue

                for position, row in enumerate(ordered_rows):
                    frame_rows = ordered_rows[: position + 1]
                    row[target_column] = self._compute_aggregate(
                        function_name,
                        arg,
                        frame_rows,
                        distinct=distinct,
                        filter_condition=filter_condition,
                    )
                continue

            raise ProgrammingError(f"Unsupported window function: {function_name}")

    @staticmethod
    def _resolve_pagination(parsed: dict[str, Any]) -> tuple[int, int | None]:
        raw_offset = parsed.get("offset")
        raw_limit = parsed.get("limit")

        if raw_offset is None:
            offset = 0
        elif isinstance(raw_offset, int):
            offset = raw_offset
        else:
            raise ProgrammingError("OFFSET must be a non-negative integer")

        if raw_limit is None:
            limit = None
        elif isinstance(raw_limit, int):
            limit = raw_limit
        else:
            raise ProgrammingError("LIMIT must be a non-negative integer")

        if offset < 0:
            raise ProgrammingError("OFFSET must be a non-negative integer")
        if limit is not None and limit < 0:
            raise ProgrammingError("LIMIT must be a non-negative integer")

        return offset, limit

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

        results: list[ExecutionResult] = []
        for query in queries:
            result = self.execute(query, _reset_subquery_cache=False)
            results.append(result)
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
                    if dname is not None and (
                        dname == col_name or dname.endswith(f".{col_name}")
                    ):
                        col_index = i
                        break
                if col_index is None:
                    raise ValueError(
                        f"ORDER BY column '{col_name}' not found in compound result"
                    )
                resolved_indexes[col_name] = col_index
            rows = self._apply_order_by(
                rows,
                order_by,
                value_getter=lambda r, col: r[resolved_indexes[col]],
            )

        compound_offset, compound_limit = self._resolve_pagination(parsed)
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
        if node.get("type") == "exists":
            return

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
                    _validate_expression_refs(
                        expression.get("expression"), column_context=column_context
                    )
                    return
                if expression_type == "unary_op":
                    _validate_expression_refs(
                        expression.get("operand"), column_context=False
                    )
                    return
                if expression_type == "binary_op":
                    _validate_expression_refs(
                        expression.get("left"), column_context=False
                    )
                    _validate_expression_refs(
                        expression.get("right"), column_context=False
                    )
                    return
                if expression_type == "function":
                    args = expression.get("args")
                    if isinstance(args, list):
                        for arg in args:
                            _validate_expression_refs(arg, column_context=False)
                    return
                if expression_type == "cast":
                    _validate_expression_refs(
                        expression.get("value"), column_context=False
                    )
                    return
                if expression_type == "case":
                    mode = str(expression.get("mode", ""))
                    if mode == "simple":
                        _validate_expression_refs(
                            expression.get("value"), column_context=False
                        )
                    for when_branch in expression.get("whens", []):
                        if not isinstance(when_branch, dict):
                            continue
                        if mode == "searched":
                            condition = when_branch.get("condition")
                            if isinstance(condition, dict):
                                SharedExecutor._validate_join_where_node(
                                    condition, validate_fn
                                )
                        else:
                            _validate_expression_refs(
                                when_branch.get("match"), column_context=False
                            )
                        _validate_expression_refs(
                            when_branch.get("result"), column_context=False
                        )
                    _validate_expression_refs(
                        expression.get("else"), column_context=False
                    )
                    return
                if expression_type in {"subquery", "exists"}:
                    return
                return
            if column_context and isinstance(expression, str) and "." in expression:
                src, col_name = expression.split(".", 1)
                validate_fn(src, col_name, "WHERE")

        column_expr = node.get("column")
        _validate_expression_refs(column_expr, column_context=True)

        value_expr = node.get("value")
        if isinstance(value_expr, dict) and value_expr.get("type") not in {
            "subquery",
            "exists",
        }:
            _validate_expression_refs(value_expr, column_context=False)
        elif isinstance(value_expr, (list, tuple)):
            for candidate in value_expr:
                _validate_expression_refs(candidate, column_context=False)

    def _matches_where(self, row: dict[str, Any], where: dict[str, Any]) -> bool:
        """Evaluate WHERE/HAVING/ON with SQL three-valued logic.

        Internally uses ``_eval_where_tv`` which returns ``True``/``False``/
        ``None`` (SQL UNKNOWN).  ``None`` is collapsed to ``False`` here so
        that all call-sites still see a plain ``bool``.
        """
        result = self._eval_where_tv(row, where)
        return result is True  # None (UNKNOWN) → False

    def _eval_where_tv(self, row: dict[str, Any], where: dict[str, Any]) -> bool | None:
        """Three-valued WHERE evaluation (True / False / None=UNKNOWN)."""
        node_type = where.get("type")
        if node_type == "exists":
            return bool(self._eval_subquery(where, outer_row=row))
        if node_type == "not":
            inner = self._eval_where_tv(row, where["operand"])
            if inner is None:
                return None  # NOT UNKNOWN = UNKNOWN
            return not inner
        if "conditions" in where:
            conditions = where["conditions"]
            conjunctions = where["conjunctions"]
            # SQL three-valued AND/OR:
            #   TRUE  AND UNKNOWN = UNKNOWN    FALSE AND UNKNOWN = FALSE
            #   TRUE  OR  UNKNOWN = TRUE       FALSE OR  UNKNOWN = UNKNOWN
            result = self._eval_where_tv(row, conditions[0])
            for idx, conj in enumerate(conjunctions):
                next_val = self._eval_where_tv(row, conditions[idx + 1])
                if conj == "AND":
                    result = _tv_and(result, next_val)
                else:  # OR
                    result = _tv_or(result, next_val)
            return result

        return self._evaluate_condition(row, where)

    def _current_outer_row(self) -> dict[str, Any] | None:
        if not self._outer_row_stack:
            return None
        return self._outer_row_stack[-1]

    def _build_scoped_row(
        self,
        row: dict[str, Any],
        *,
        headers: list[str] | None = None,
        source_refs: set[str] | None = None,
    ) -> dict[str, Any]:
        scoped_row = dict(row)

        if headers is not None and source_refs:
            for source_ref in source_refs:
                for header in headers:
                    qualified_name = f"{source_ref}.{header}"
                    if qualified_name not in scoped_row:
                        scoped_row[qualified_name] = row.get(header)

        outer_row = self._current_outer_row()
        if outer_row is not None:
            for key, value in outer_row.items():
                if key not in scoped_row:
                    scoped_row[key] = value

        return scoped_row

    def _row_from_values(
        self, headers: list[str], row_values: list[Any]
    ) -> dict[str, Any]:
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

    def _resolve_join_column(
        self, row: dict[str, Any], col_spec: dict[str, Any]
    ) -> Any:
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

    def _matches_join_on_condition(
        self,
        left_ns: dict[str, dict[str, Any]],
        right_ns: dict[str, dict[str, Any]],
        on_condition: dict[str, Any] | None,
    ) -> bool:
        if on_condition is None:
            return True

        combined_row: dict[str, dict[str, Any]] = {}
        combined_row.update(left_ns)
        combined_row.update(right_ns)
        flattened = self._build_scoped_row(self._flatten_join_row(combined_row))
        return self._matches_where(flattened, on_condition)

    def _join_two_sources(
        self,
        left_rows: list[dict[str, dict[str, Any]]],
        left_headers_map: dict[str, set[str]],
        right_data: TableData,
        right_source: dict[str, Any],
        join_type: str,
        on_condition: dict[str, Any] | None,
    ) -> tuple[list[dict[str, dict[str, Any]]], dict[str, set[str]]]:
        right_headers = list(right_data.headers)
        right_sources = {str(right_source["table"]), str(right_source["ref"])}

        right_rows = [
            self._build_source_row(right_source, right_headers, right_row_values)
            for right_row_values in right_data.rows
        ]
        join_type_upper = join_type.upper()

        joined_rows: list[dict[str, dict[str, Any]]] = []
        right_null_values = [None for _ in right_headers]
        right_null_ns = self._build_source_row(
            right_source, right_headers, right_null_values
        )
        left_null_ns = {
            source: {column: None for column in columns}
            for source, columns in left_headers_map.items()
        }

        if join_type_upper == "CROSS":
            for left_ns in left_rows:
                for right_ns in right_rows:
                    combined_row = {}
                    combined_row.update(left_ns)
                    combined_row.update(right_ns)
                    joined_rows.append(combined_row)
        else:
            matched_right_indices: set[int] = set()

            for left_ns in left_rows:
                left_matched = False
                for right_index, right_ns in enumerate(right_rows):
                    if not self._matches_join_on_condition(
                        left_ns, right_ns, on_condition
                    ):
                        continue

                    left_matched = True
                    matched_right_indices.add(right_index)
                    combined_row = {}
                    combined_row.update(left_ns)
                    combined_row.update(right_ns)
                    joined_rows.append(combined_row)

                if not left_matched and join_type_upper in {"LEFT", "FULL"}:
                    combined_row = {}
                    combined_row.update(left_ns)
                    combined_row.update(right_null_ns)
                    joined_rows.append(combined_row)

            if join_type_upper in {"RIGHT", "FULL"}:
                for right_index, right_ns in enumerate(right_rows):
                    if right_index in matched_right_indices:
                        continue
                    combined_row = {}
                    combined_row.update(left_null_ns)
                    combined_row.update(right_ns)
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
        left_data: TableData,
    ) -> ExecutionResult:
        joins = parsed.get("joins") or []
        from_source = parsed["from"]
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
            resolved_right_table, right_data = self._resolve_table_data(
                str(right_source["table"])
            )
            if resolved_right_table is None or right_data is None:
                available = self._available_table_names()
                msg = f"Sheet '{right_source['table']}' not found in Excel."
                if available:
                    msg += f" Available sheets: {available}"
                raise ValueError(msg)

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
            if join_on is not None:
                SharedExecutor._validate_join_where_node(join_on, _validate_column_ref)

            known_sources.update({right_ref, right_table_name})
            join_inputs.append((join_spec, right_data))

        # Validate SELECT columns
        def _validate_join_select_expression(expression: Any) -> None:
            if expression is None:
                return
            if isinstance(expression, str):
                if "." not in expression:
                    raise ValueError(
                        "JOIN queries require qualified column names in SELECT"
                    )
                source, name = expression.split(".", 1)
                _validate_column_ref(source, name, "SELECT")
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
                    if arg != "*":
                        if "." not in arg:
                            raise ValueError(
                                "Aggregate arguments in JOIN queries must be qualified column names or *"
                            )
                        source, name = arg.split(".", 1)
                        _validate_column_ref(source, name, "SELECT")

                    filter_clause = expression.get("filter")
                    if isinstance(filter_clause, dict):
                        SharedExecutor._validate_join_where_node(
                            filter_clause, _validate_column_ref
                        )
                    return
                if expression_type == "window_function":
                    args = expression.get("args")
                    if isinstance(args, list):
                        for argument in args:
                            if isinstance(argument, str):
                                if argument == "*":
                                    continue
                                if "." not in argument:
                                    raise ValueError(
                                        "Window function arguments in JOIN queries must be qualified column names or *"
                                    )
                                source, name = argument.split(".", 1)
                                _validate_column_ref(source, name, "SELECT")
                            else:
                                _validate_join_select_expression(argument)

                    partition_by = expression.get("partition_by")
                    if isinstance(partition_by, list):
                        for partition_expression in partition_by:
                            _validate_join_select_expression(partition_expression)

                    order_by_items = expression.get("order_by")
                    if isinstance(order_by_items, list):
                        for order_item in order_by_items:
                            if not isinstance(order_item, dict):
                                continue
                            order_expression = order_item.get("__expression__")
                            if order_expression is not None:
                                _validate_join_select_expression(order_expression)
                                continue

                            order_column = order_item.get("column")
                            if isinstance(order_column, str):
                                if order_column.startswith("__expr__:"):
                                    continue
                                if "." not in order_column:
                                    raise ValueError(
                                        "Window function ORDER BY in JOIN queries requires qualified column names"
                                    )
                                source, name = order_column.split(".", 1)
                                _validate_column_ref(source, name, "SELECT")

                    filter_clause = expression.get("filter")
                    if isinstance(filter_clause, dict):
                        SharedExecutor._validate_join_where_node(
                            filter_clause, _validate_column_ref
                        )
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
                if expression_type == "function":
                    args = expression.get("args")
                    if isinstance(args, list):
                        for arg in args:
                            _validate_join_select_expression(arg)
                    return
                if expression_type == "cast":
                    _validate_join_select_expression(expression.get("value"))
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
                                SharedExecutor._validate_join_where_node(
                                    condition, _validate_column_ref
                                )
                        else:
                            _validate_join_select_expression(when_branch.get("match"))
                        _validate_join_select_expression(when_branch.get("result"))
                    _validate_join_select_expression(expression.get("else"))
                    return
                if expression_type == "subquery":
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
        for join_spec, right_data in join_inputs:
            right_source = join_spec["source"]
            join_on = join_spec.get("on")
            left_rows, left_headers_map = self._join_two_sources(
                left_rows=left_rows,
                left_headers_map=left_headers_map,
                right_data=right_data,
                right_source=right_source,
                join_type=str(join_spec["type"]),
                on_condition=join_on,
            )

        joined_rows_flat = [
            self._build_scoped_row(self._flatten_join_row(row)) for row in left_rows
        ]

        where = parsed.get("where")
        if where:
            joined_rows_flat = [
                row for row in joined_rows_flat if self._matches_where(row, where)
            ]

        window_columns = self._apply_window_functions(
            joined_rows_flat, parsed["columns"], order_by
        )

        columns = parsed["columns"]
        aggregate_query = (
            any(self._is_aggregate_column(self._unwrap_alias(col)) for col in columns)
            if columns != ["*"]
            else False
        )
        group_by = parsed.get("group_by")
        having = parsed.get("having")

        if columns == ["*"] and (aggregate_query or group_by is not None):
            raise ValueError(
                "SELECT * is not supported with GROUP BY or aggregate functions"
            )

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

        distinct = bool(parsed.get("distinct", False))
        selected_columns: list[str] = []
        output_names: list[str] = []
        if columns == ["*"]:
            for source_ref, ordered_headers in source_headers_ordered:
                for col_name in ordered_headers:
                    selected_columns.append(f"{source_ref}.{col_name}")
                    output_names.append(f"{source_ref}.{col_name}")

            rows_for_output = joined_rows_flat
            if distinct:
                self._validate_distinct_order_by_columns(order_by, selected_columns)
                rows_for_output = self._dedupe_projected_rows(
                    rows_for_output,
                    selected_columns,
                )

            if order_by:
                order_expression_columns = self._materialize_order_expression_columns(
                    rows_for_output,
                    order_by,
                )
                available_cols: set[str] = set()
                for source_ref, ordered_headers in source_headers_ordered:
                    for col_name in ordered_headers:
                        available_cols.add(f"{source_ref}.{col_name}")
                available_cols.update(order_expression_columns)
                available_cols.update(window_columns)
                rows_for_output = self._apply_order_by(
                    rows_for_output,
                    order_by,
                    value_getter=lambda r, col: r.get(col),
                    available_columns=available_cols,
                )

            offset, limit = self._resolve_pagination(parsed)
            if offset:
                rows_for_output = rows_for_output[offset:]
            if limit is not None:
                rows_for_output = rows_for_output[:limit]

            rows_out = [
                tuple(row.get(column_name) for column_name in selected_columns)
                for row in rows_for_output
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

            if distinct:
                self._validate_distinct_order_by_columns(order_by, selected_columns)
                projected_rows = self._dedupe_projected_rows(
                    projected_rows,
                    selected_columns,
                )

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
                available_columns.update(window_columns)
                projected_rows = self._apply_order_by(
                    projected_rows,
                    order_by,
                    value_getter=lambda r, col: r.get(col),
                    available_columns=available_columns,
                )

            offset, limit = self._resolve_pagination(parsed)
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

        group_by: list[Any] | None = parsed.get("group_by")
        having = parsed.get("having")
        aggregate_query = any(self._is_aggregate_column(col) for col in columns)

        if columns == ["*"] and (aggregate_query or group_by):
            raise ValueError(
                "SELECT * is not supported with GROUP BY or aggregate functions"
            )

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

        window_columns = self._apply_window_functions(rows, columns, order_by)

        projected_rows: list[dict[str, Any]]
        if needs_expression_projection and columns != ["*"]:
            projected_rows = []
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
        else:
            projected_rows = rows

        distinct = parsed.get("distinct", False)
        if distinct:
            self._validate_distinct_order_by_columns(order_by, selected_columns)
            projected_rows = self._dedupe_projected_rows(
                projected_rows, selected_columns
            )

        if order_by:
            expression_columns = self._materialize_order_expression_columns(
                projected_rows,
                order_by,
            )
            available_columns = set(headers)
            available_columns.update(selected_columns)
            available_columns.update(expression_columns)
            available_columns.update(window_columns)
            projected_rows = self._apply_order_by(
                projected_rows,
                order_by,
                value_getter=lambda r, col: r.get(col),
                available_columns=available_columns,
            )

        offset, limit = self._resolve_pagination(parsed)
        if offset:
            projected_rows = projected_rows[offset:]
        if limit is not None:
            projected_rows = projected_rows[:limit]

        rows_out = [
            tuple(projected_row.get(col) for col in selected_columns)
            for projected_row in projected_rows
        ]

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
            mode = str(value.get("mode", "set"))
            if mode != "set" or bool(value.get("correlated", False)):
                return

            subquery_result = self.execute(
                value["query"],
                _reset_subquery_cache=False,
            )
            node["value"] = tuple(row[0] for row in subquery_result.rows if row)

    def _eval_subquery(
        self,
        node: dict[str, Any],
        *,
        outer_row: dict[str, Any] | None = None,
    ) -> Any:
        node_type = str(node.get("type"))
        if node_type not in {"subquery", "exists"}:
            raise ProgrammingError(f"Unsupported subquery node type: {node_type}")

        correlated = bool(node.get("correlated", False))
        cache_key: int | None = None
        if not correlated:
            cache_key = id(node)
            cached = self._subquery_cache.get(cache_key)
            if cache_key in self._subquery_cache:
                return cached

        query = node.get("query")
        if not isinstance(query, dict):
            raise ProgrammingError("Invalid subquery node: missing parsed query")

        pushed_outer_row = False
        if correlated and outer_row is not None:
            self._outer_row_stack.append(outer_row)
            pushed_outer_row = True
        try:
            result = self.execute(query, _reset_subquery_cache=False)
        finally:
            if pushed_outer_row:
                self._outer_row_stack.pop()

        evaluated: Any
        if node_type == "exists":
            evaluated = bool(result.rows)
        else:
            mode = str(node.get("mode", "set"))
            if mode == "set":
                evaluated = tuple(row[0] for row in result.rows if row)
            elif mode == "scalar":
                if len(result.rows) > 1:
                    raise ProgrammingError("Scalar subquery returned more than one row")
                if not result.rows:
                    evaluated = None
                else:
                    first_row = result.rows[0]
                    evaluated = first_row[0] if first_row else None
            else:
                raise ProgrammingError(f"Unsupported subquery mode: {mode}")

        if cache_key is not None:
            self._subquery_cache[cache_key] = evaluated

        return evaluated

    def _execute_aggregate_select(
        self,
        action: str,
        parsed: dict[str, Any],
        headers: list[str],
        rows: list[dict[str, Any]],
        columns: list[Any],
        group_by: list[Any] | None,
        having: dict[str, Any] | None,
    ) -> ExecutionResult:
        group_entries: list[tuple[str, Any]] = []
        if group_by is not None:
            for group_expression in group_by:
                group_key = self._source_key(group_expression)
                referenced_columns = self._collect_expression_column_refs(
                    group_expression
                )
                missing_group_columns = sorted(
                    column_name
                    for column_name in referenced_columns
                    if column_name not in headers
                )
                if missing_group_columns:
                    raise ValueError(
                        "Unknown column(s): "
                        f"{', '.join(missing_group_columns)}. Available columns: {headers}"
                    )
                group_entries.append((group_key, group_expression))
        group_by_keys = {key for key, _ in group_entries}

        def _operand_reference(operand: Any) -> str | None:
            if isinstance(operand, str):
                return operand
            if isinstance(operand, dict) and operand.get("type") != "subquery":
                return self._source_key(operand)
            return None

        def _iter_having_operand_refs(node: Any) -> Iterator[str]:
            if not isinstance(node, dict):
                return

            node_type = node.get("type")
            if node_type == "not":
                yield from _iter_having_operand_refs(node.get("operand"))
                return
            if node_type == "exists":
                return

            if "conditions" in node:
                for child in node.get("conditions", []):
                    yield from _iter_having_operand_refs(child)
                return

            for candidate in (node.get("column"), node.get("value")):
                if isinstance(candidate, (list, tuple)):
                    for item in candidate:
                        ref = _operand_reference(item)
                        if ref is not None:
                            yield ref
                    continue
                ref = _operand_reference(candidate)
                if ref is not None:
                    yield ref

        output_columns: list[str] = []
        output_sources: list[str] = []
        required_aggregates: dict[
            str, tuple[str, str, bool, dict[str, Any] | None]
        ] = {}
        for column in columns:
            inner = self._unwrap_alias(column)
            if self._is_aggregate_column(inner):
                aggregate_column = inner
                raw_arg = str(aggregate_column.get("arg"))
                arg = self._normalize_single_source_aggregate_arg(raw_arg, headers)
                if arg != "*" and arg not in headers:
                    raise ValueError(
                        f"Unknown column: {raw_arg}. Available columns: {headers}"
                    )
                label = self._aggregate_label(aggregate_column)
                filter_clause = aggregate_column.get("filter")
                filter_condition = (
                    filter_clause if isinstance(filter_clause, dict) else None
                )
                required_aggregates[label] = (
                    str(aggregate_column.get("func")),
                    arg,
                    bool(aggregate_column.get("distinct")),
                    filter_condition,
                )
                output_columns.append(self._output_name(column))
                output_sources.append(label)
                continue

            column_name = self._source_key(column)
            referenced_columns = self._collect_expression_column_refs(inner)
            missing_selected_columns = sorted(
                ref for ref in referenced_columns if ref not in headers
            )
            if missing_selected_columns:
                raise ValueError(
                    f"Unknown column(s): {', '.join(missing_selected_columns)}. Available columns: {headers}"
                )
            if group_by is None:
                raise ValueError(
                    "Non-aggregate columns in aggregate queries require GROUP BY"
                )
            if column_name not in group_by_keys:
                raise ValueError(
                    f"Selected column '{column_name}' must appear in GROUP BY"
                )
            output_columns.append(self._output_name(column))
            output_sources.append(column_name)

        if having is not None:
            for ref in _iter_having_operand_refs(having):
                aggregate_spec = self._aggregate_spec_from_label(ref)
                if aggregate_spec is None:
                    continue
                func, arg, distinct, filter_condition = aggregate_spec
                normalized_arg = self._normalize_single_source_aggregate_arg(
                    arg, headers
                )
                if normalized_arg != "*" and normalized_arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                required_aggregates[ref] = (
                    func,
                    normalized_arg,
                    distinct,
                    filter_condition,
                )

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
                if ref.startswith("__expr__:"):
                    continue
                aggregate_spec = self._aggregate_spec_from_label(ref)
                if aggregate_spec is None:
                    continue
                func, arg, distinct, filter_condition = aggregate_spec
                normalized_arg = self._normalize_single_source_aggregate_arg(
                    arg, headers
                )
                if normalized_arg != "*" and normalized_arg not in headers:
                    raise ValueError(
                        f"Unknown column: {arg}. Available columns: {headers}"
                    )
                required_aggregates[ref] = (
                    func,
                    normalized_arg,
                    distinct,
                    filter_condition,
                )

        if having is not None:
            for column_ref in self._collect_where_column_refs(having):
                if self._aggregate_spec_from_label(column_ref) is not None:
                    continue
                if group_by is None or column_ref not in group_by_keys:
                    raise ValueError(
                        f"Column '{column_ref}' in HAVING must be a GROUP BY column or aggregate function"
                    )

        grouped_rows: list[dict[str, Any]] = []
        if group_entries:
            groups: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
            for row in rows:
                group_values: list[Any] = []
                for _, group_expression in group_entries:
                    if isinstance(group_expression, dict):
                        group_values.append(
                            self._eval_expression(
                                group_expression,
                                row,
                                lambda col_name: row.get(col_name),
                            )
                        )
                    else:
                        group_values.append(row.get(str(group_expression)))
                groups.setdefault(tuple(group_values), []).append(row)

            for grouped_key, group_values in groups.items():
                context_row: dict[str, Any] = (
                    dict(group_values[0]) if group_values else {}
                )
                for (group_key, _), group_value in zip(group_entries, grouped_key):
                    context_row[group_key] = group_value
                for label, (
                    func,
                    arg,
                    distinct,
                    filter_condition,
                ) in required_aggregates.items():
                    context_row[label] = self._compute_aggregate(
                        func,
                        arg,
                        group_values,
                        distinct=distinct,
                        filter_condition=filter_condition,
                    )
                if having and not self._matches_where(context_row, having):
                    continue
                grouped_rows.append(context_row)
        else:
            context_row_single: dict[str, Any] = {}
            for label, (
                func,
                arg,
                distinct,
                filter_condition,
            ) in required_aggregates.items():
                context_row_single[label] = self._compute_aggregate(
                    func,
                    arg,
                    rows,
                    distinct=distinct,
                    filter_condition=filter_condition,
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
            available_order_columns.update(group_by_keys)
            available_order_columns.update(required_aggregates.keys())
            available_order_columns.update(order_expression_columns)
            grouped_rows = self._apply_order_by(
                grouped_rows,
                order_by_clause,
                value_getter=lambda r, col: r.get(col),
                available_columns=available_order_columns,
            )

        offset, limit = self._resolve_pagination(parsed)
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
        return self._expression_to_sql(aggregate)

    def _aggregate_spec_from_label(
        self, expression: str
    ) -> tuple[str, str, bool, dict[str, Any] | None] | None:
        expression_text = expression.strip()
        if not expression_text or expression_text.startswith("__expr__:"):
            return None

        try:
            parsed_expression = _parse_column_expression(
                expression_text,
                allow_wildcard=False,
                allow_aggregates=True,
                allow_subqueries=False,
            )
        except ValueError as exc:
            if "DISTINCT is only supported with COUNT" in str(exc):
                raise
            return None
        if (
            not isinstance(parsed_expression, dict)
            or parsed_expression.get("type") != "aggregate"
        ):
            return None

        func = str(parsed_expression.get("func", "")).upper()
        arg = str(parsed_expression.get("arg", "")).strip()
        if not arg:
            return None

        filter_clause = parsed_expression.get("filter")
        filter_condition = filter_clause if isinstance(filter_clause, dict) else None
        return func, arg, bool(parsed_expression.get("distinct")), filter_condition

    def _compute_aggregate(
        self,
        func: str,
        arg: str,
        rows: list[dict[str, Any]],
        distinct: bool = False,
        filter_condition: dict[str, Any] | None = None,
    ) -> Any:
        applicable_rows = rows
        if filter_condition is not None:
            applicable_rows = [
                row for row in rows if self._matches_where(row, filter_condition)
            ]

        aggregate = func.upper()
        if aggregate == "COUNT":
            if distinct:
                if arg == "*":
                    raise ValueError("COUNT(DISTINCT *) is not supported")
                return len(
                    {
                        row.get(arg)
                        for row in applicable_rows
                        if row.get(arg) is not None
                    }
                )
            if arg == "*":
                return len(applicable_rows)
            return sum(1 for row in applicable_rows if row.get(arg) is not None)

        values: list[_SupportsOrder] = []
        for row in applicable_rows:
            value = row.get(arg)
            if value is not None:
                values.append(cast(_SupportsOrder, value))
        if not values:
            return None

        numeric_values: list[float] = []
        for value in values:
            numeric = self._to_number(value)
            if numeric is not None:
                numeric_values.append(numeric)

        if aggregate == "SUM":
            if not numeric_values:
                return None
            return sum(numeric_values)
        if aggregate == "AVG":
            if not numeric_values:
                return None
            return sum(numeric_values) / len(numeric_values)
        if aggregate == "MIN":
            return min(values)
        if aggregate == "MAX":
            return max(values)
        raise ValueError(f"Unsupported aggregate function: {func}")

    def _call_function(self, name: str, args: list[Any]) -> Any:
        normalized_name = name.upper()
        function_spec = _SCALAR_FUNCTIONS.get(normalized_name)
        if function_spec is None:
            raise ProgrammingError(f"Unsupported scalar function: {name}")

        min_args, max_args, function_handler = function_spec
        if len(args) < min_args:
            raise ProgrammingError(
                f"{normalized_name} expects at least {min_args} argument(s), got {len(args)}"
            )
        if max_args is not None and len(args) > max_args:
            raise ProgrammingError(
                f"{normalized_name} expects at most {max_args} argument(s), got {len(args)}"
            )

        try:
            return function_handler(args)
        except (TypeError, ValueError) as exc:
            raise ProgrammingError(
                f"Invalid arguments for {normalized_name}: {exc}"
            ) from exc

    def _eval_cast(self, value: Any, target_type: str) -> Any:
        if value is None:
            return None

        normalized_target = target_type.upper()
        if normalized_target in {"INTEGER", "INT"}:
            numeric_value = self._to_number(value)
            if numeric_value is None:
                raise ProgrammingError(
                    f"Cannot cast value {value!r} to {normalized_target}"
                )
            return int(numeric_value)

        if normalized_target in {"REAL", "FLOAT", "NUMERIC"}:
            numeric_value = self._to_number(value)
            if numeric_value is None:
                raise ProgrammingError(
                    f"Cannot cast value {value!r} to {normalized_target}"
                )
            return float(numeric_value)

        if normalized_target == "TEXT":
            if isinstance(value, (date, datetime)):
                return value.isoformat()
            return str(value)

        if normalized_target == "DATE":
            if isinstance(value, datetime):
                return value.date()
            if isinstance(value, date):
                return value
            if isinstance(value, str):
                stripped = value.strip()
                if not stripped:
                    raise ProgrammingError("Cannot cast empty string to DATE")
                try:
                    return date.fromisoformat(stripped)
                except ValueError:
                    normalized_datetime = stripped
                    if normalized_datetime.endswith("Z"):
                        normalized_datetime = normalized_datetime[:-1] + "+00:00"
                    try:
                        return datetime.fromisoformat(normalized_datetime).date()
                    except ValueError as exc:
                        raise ProgrammingError(
                            f"Cannot cast value {value!r} to DATE"
                        ) from exc
            raise ProgrammingError(f"Cannot cast value {value!r} to DATE")

        if normalized_target == "DATETIME":
            if isinstance(value, datetime):
                return value.replace(tzinfo=None) if value.tzinfo is not None else value
            if isinstance(value, date):
                return datetime.combine(value, time.min)
            if isinstance(value, str):
                stripped = value.strip()
                if not stripped:
                    raise ProgrammingError("Cannot cast empty string to DATETIME")
                normalized_datetime = stripped
                if normalized_datetime.endswith("Z"):
                    normalized_datetime = normalized_datetime[:-1] + "+00:00"
                try:
                    parsed_datetime = datetime.fromisoformat(normalized_datetime)
                except ValueError as exc:
                    raise ProgrammingError(
                        f"Cannot cast value {value!r} to DATETIME"
                    ) from exc
                return (
                    parsed_datetime.replace(tzinfo=None)
                    if parsed_datetime.tzinfo is not None
                    else parsed_datetime
                )
            raise ProgrammingError(f"Cannot cast value {value!r} to DATETIME")

        if normalized_target == "BOOLEAN":
            if isinstance(value, bool):
                return value
            if isinstance(value, (int, float)) and value in {0, 1}:
                return bool(value)
            if isinstance(value, str):
                normalized_boolean = value.strip().lower()
                if normalized_boolean in {"1", "true", "yes"}:
                    return True
                if normalized_boolean in {"0", "false", "no"}:
                    return False
            raise ProgrammingError(f"Cannot cast value {value!r} to BOOLEAN")

        raise ProgrammingError(f"Unsupported CAST target type: {target_type}")

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

        if expr_type == "subquery":
            return self._eval_subquery(expr, outer_row=row)

        if expr_type == "unary_op":
            if expr.get("op") != "-":
                raise ProgrammingError(f"Unsupported unary operator: {expr.get('op')}")
            operand_value = self._eval_expression(
                expr.get("operand"), row, resolve_column
            )
            if operand_value is None:
                return None
            numeric_operand = self._to_number(operand_value)
            if numeric_operand is None:
                raise ProgrammingError(
                    "Arithmetic expression requires numeric operands"
                )
            return -numeric_operand

        if expr_type == "binary_op":
            operator = str(expr.get("op"))
            left_value = self._eval_expression(expr.get("left"), row, resolve_column)
            right_value = self._eval_expression(expr.get("right"), row, resolve_column)
            if left_value is None or right_value is None:
                return None

            if operator == "||":
                return f"{left_value}{right_value}"

            left_number = self._to_number(left_value)
            right_number = self._to_number(right_value)
            if left_number is None or right_number is None:
                raise ProgrammingError(
                    "Arithmetic expression requires numeric operands"
                )

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

        if expr_type == "function":
            args = expr.get("args")
            args_list = args if isinstance(args, list) else []
            evaluated_args = [
                self._eval_expression(argument, row, resolve_column)
                for argument in args_list
            ]
            return self._call_function(str(expr.get("name", "")), evaluated_args)

        if expr_type == "cast":
            value = self._eval_expression(expr.get("value"), row, resolve_column)
            target_type = str(expr.get("target_type", ""))
            return self._eval_cast(value, target_type)

        if expr_type == "window_function":
            return resolve_column(self._source_key(expr))

        if expr_type == "aggregate":
            raise ProgrammingError(
                "Aggregate expressions are not supported in row-level arithmetic"
            )

        if expr_type == "case":
            mode = expr.get("mode", "searched")
            if mode == "searched":
                for when_branch in expr.get("whens", []):
                    if not isinstance(when_branch, dict):
                        continue
                    condition = when_branch.get("condition")
                    if isinstance(condition, dict) and self._matches_where(
                        row, condition
                    ):
                        return self._eval_expression(
                            when_branch["result"], row, resolve_column
                        )
            else:  # simple mode
                case_value = self._eval_expression(
                    expr.get("value"), row, resolve_column
                )
                for when_branch in expr.get("whens", []):
                    if not isinstance(when_branch, dict):
                        continue
                    match_value = self._eval_expression(
                        when_branch["match"], row, resolve_column
                    )
                    if case_value is not None and match_value is not None:
                        left, right = self._coerce_for_compare(case_value, match_value)
                        if left == right:
                            return self._eval_expression(
                                when_branch["result"], row, resolve_column
                            )
            # No WHEN matched — evaluate ELSE or return None
            else_expr = expr.get("else")
            if else_expr is not None:
                return self._eval_expression(else_expr, row, resolve_column)
            return None

        raise ProgrammingError(f"Unsupported expression type: {expr_type}")

    def _evaluate_condition(
        self, row: dict[str, Any], condition: dict[str, Any]
    ) -> bool | None:
        if condition.get("type") == "exists":
            return bool(self._eval_subquery(condition, outer_row=row))

        column = condition["column"]
        operator = condition["operator"]
        value = condition["value"]

        def _resolve_operand(operand: Any, *, as_column: bool) -> Any:
            if isinstance(operand, dict):
                if operand.get("type") in {"subquery", "exists"}:
                    return self._eval_subquery(operand, outer_row=row)
                return self._eval_expression(
                    operand, row, lambda col_name: row.get(col_name)
                )
            if as_column:
                return row.get(str(operand))
            if type(operand) is str and operand in row:
                return row.get(operand)
            return operand

        row_value = _resolve_operand(column, as_column=True)

        if operator == "IS" and value is None:
            return row_value is None
        if operator == "IS NOT" and value is None:
            return row_value is not None
        if operator == "IN":
            if row_value is None:
                return None  # SQL UNKNOWN
            resolved_candidates = _resolve_operand(value, as_column=False)
            if resolved_candidates is None:
                candidates: tuple[Any, ...] = ()
            elif isinstance(resolved_candidates, tuple):
                candidates = resolved_candidates
            elif isinstance(resolved_candidates, list):
                candidates = tuple(resolved_candidates)
            else:
                candidates = (resolved_candidates,)

            has_null = False
            for candidate in candidates:
                candidate_value = _resolve_operand(candidate, as_column=False)
                if candidate_value is None:
                    has_null = True
                    continue
                left, right = self._coerce_for_compare(row_value, candidate_value)
                if left == right:
                    return True
            # SQL: x IN (..., NULL, ...) is UNKNOWN when no match
            return None if has_null else False
        if operator == "NOT IN":
            if row_value is None:
                return None  # SQL UNKNOWN
            resolved_candidates = _resolve_operand(value, as_column=False)
            if resolved_candidates is None:
                candidates = ()
            elif isinstance(resolved_candidates, tuple):
                candidates = resolved_candidates
            elif isinstance(resolved_candidates, list):
                candidates = tuple(resolved_candidates)
            else:
                candidates = (resolved_candidates,)

            has_null = False
            for candidate in candidates:
                candidate_value = _resolve_operand(candidate, as_column=False)
                if candidate_value is None:
                    has_null = True
                    continue
                left, right = self._coerce_for_compare(row_value, candidate_value)
                if left == right:
                    return False
            return None if has_null else True
        if operator == "BETWEEN":
            if row_value is None:
                return None  # SQL UNKNOWN
            low, high = value
            low = _resolve_operand(low, as_column=False)
            high = _resolve_operand(high, as_column=False)
            if low is None or high is None:
                return None  # SQL UNKNOWN
            left_low, right_low = self._coerce_for_compare(row_value, low)
            left_high, right_high = self._coerce_for_compare(row_value, high)
            return bool(left_low >= right_low and left_high <= right_high)
        if operator == "NOT BETWEEN":
            if row_value is None:
                return None  # SQL UNKNOWN
            low, high = value
            low = _resolve_operand(low, as_column=False)
            high = _resolve_operand(high, as_column=False)
            if low is None or high is None:
                return None  # SQL UNKNOWN
            left_low, right_low = self._coerce_for_compare(row_value, low)
            left_high, right_high = self._coerce_for_compare(row_value, high)
            return not bool(left_low >= right_low and left_high <= right_high)
        if operator in {"LIKE", "NOT LIKE", "ILIKE", "NOT ILIKE"}:
            if row_value is None:
                return None  # SQL UNKNOWN
            value = _resolve_operand(value, as_column=False)
            if value is None:
                return None  # SQL UNKNOWN — NULL pattern
            if not isinstance(value, str):
                raise NotImplementedError("Unsupported LIKE pattern type")
            escape_value = condition.get("escape")
            if escape_value is not None:
                if not isinstance(escape_value, str) or len(escape_value) != 1:
                    raise ValueError("ESCAPE requires a single character")

            row_text = str(row_value)
            pattern = value
            if operator in {"ILIKE", "NOT ILIKE"}:
                row_text = row_text.lower()
                pattern = pattern.lower()

            regex = _build_like_regex(pattern, escape_value)
            is_match = bool(re.match(regex, row_text))
            if operator in {"NOT LIKE", "NOT ILIKE"}:
                return not is_match
            return is_match

        value = _resolve_operand(value, as_column=False)

        if row_value is None or value is None:
            return None  # SQL UNKNOWN

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
        if isinstance(left, bool) and isinstance(right, bool):
            return int(left), int(right)

        left_temporal = self._coerce_temporal_value(left)
        right_temporal = self._coerce_temporal_value(right)
        if left_temporal is not None and right_temporal is not None:
            return left_temporal, right_temporal

        left_num = self._to_number(left)
        right_num = self._to_number(right)
        if left_num is not None and right_num is not None:
            return left_num, right_num
        return str(left if left is not None else ""), str(
            right if right is not None else ""
        )

    def _sort_key(self, value: Any) -> tuple[int, Any]:
        if value is None:
            return (1, (0, ""))

        if isinstance(value, bool):
            return (0, (1, int(value)))

        temporal = self._coerce_temporal_value(value)
        if temporal is not None:
            return (0, (2, temporal))

        numeric = self._to_number(value)
        if numeric is not None:
            return (0, (0, numeric))

        if isinstance(value, str):
            return (0, (3, value))

        return (0, (4, str(value)))

    @staticmethod
    def _coerce_temporal_value(value: Any) -> datetime | None:
        if isinstance(value, datetime):
            return value.replace(tzinfo=None) if value.tzinfo is not None else value
        if isinstance(value, date):
            return datetime.combine(value, time.min)
        if isinstance(value, str):
            return SharedExecutor._parse_datetime_string(value)
        return None

    @staticmethod
    def _parse_datetime_string(value: str) -> datetime | None:
        text = value.strip()
        if not text:
            return None

        normalized = text
        if normalized.endswith("Z"):
            normalized = normalized[:-1] + "+00:00"

        try:
            parsed_datetime = datetime.fromisoformat(normalized)
        except ValueError:
            parsed_datetime = None

        if parsed_datetime is not None:
            return (
                parsed_datetime.replace(tzinfo=None)
                if parsed_datetime.tzinfo is not None
                else parsed_datetime
            )

        try:
            parsed_date = date.fromisoformat(text)
        except ValueError:
            parsed_date = None

        if parsed_date is not None:
            return datetime.combine(parsed_date, time.min)

        formats = (
            "%Y/%m/%d",
            "%Y/%m/%d %H:%M:%S",
            "%Y/%m/%d %H:%M",
            "%m/%d/%Y",
            "%m/%d/%Y %H:%M:%S",
            "%m/%d/%Y %H:%M",
            "%d/%m/%Y",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y %H:%M",
        )
        for date_format in formats:
            try:
                return datetime.strptime(text, date_format)
            except ValueError:
                continue

        return None

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

    def _resolve_cte_name(self, requested_name: str) -> str | None:
        lowered = requested_name.lower()
        for name in self._cte_tables:
            if name.lower() == lowered:
                return name
        return None

    def _resolve_table_data(
        self, requested_name: str
    ) -> tuple[str | None, TableData | None]:
        cte_name = self._resolve_cte_name(requested_name)
        if cte_name is not None:
            return cte_name, self._cte_tables.get(cte_name)

        resolved_sheet = self._resolve_sheet_name(requested_name)
        if resolved_sheet is None:
            return None, None
        return resolved_sheet, self.backend.read_sheet(resolved_sheet)

    def _available_table_names(self) -> list[str]:
        names = list(self.backend.list_sheets())
        names.extend(self._cte_tables.keys())
        return names
