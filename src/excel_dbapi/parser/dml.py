from __future__ import annotations

from typing import Any, Dict, List, Optional

from ._constants import _is_placeholder
from .tokenizer import (
    _count_unquoted_placeholders,
    _parse_column_identifier,
    _parse_table_identifier,
    _parse_value,
    _split_csv,
    _tokenize,
)
from .expressions import (
    _annotate_column_tables,
    _bind_expression_values,
    _bind_where_conditions,
    _expression_values_to_bind,
    _parse_column_expression,
    _where_values_to_bind,
)
from .where import _parse_where_expression
from .select import _bind_params, _parse_columns, _parse_select


def _split_insert_on_conflict_clause(remainder: str) -> tuple[str, Optional[str]]:
    tokens = _tokenize(remainder.strip())
    if not tokens:
        return "", None

    depth = 0
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if token == "(":
            depth += 1
        elif token == ")":
            if depth > 0:
                depth -= 1
        elif (
            depth == 0
            and token.upper() == "ON"
            and index + 1 < len(tokens)
            and tokens[index + 1].upper() == "CONFLICT"
        ):
            before = " ".join(tokens[:index]).strip()
            on_conflict = " ".join(tokens[index:]).strip()
            return before, on_conflict
        index += 1

    return remainder.strip(), None


def _parse_assignment_expression(
    value_text: str,
    *,
    annotate_tables: bool,
    error_message: str,
) -> Any:
    stripped = value_text.strip()
    if not stripped:
        raise ValueError(error_message)
    parsed = _parse_column_expression(
        stripped,
        allow_wildcard=False,
        allow_aggregates=False,
    )
    if annotate_tables:
        _annotate_column_tables(parsed)
    return parsed


def _parse_upsert_assignment_value(value_text: str) -> Any:
    return _parse_assignment_expression(
        value_text,
        annotate_tables=True,
        error_message="Invalid ON CONFLICT clause format",
    )


def _parse_on_conflict_clause(
    clause: str,
    query: str,
    params: Optional[tuple[Any, ...]],
) -> Dict[str, Any]:
    tokens = _tokenize(clause.strip())
    if len(tokens) < 5:
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
    if tokens[0].upper() != "ON" or tokens[1].upper() != "CONFLICT":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    index = 2
    if index >= len(tokens) or tokens[index] != "(":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    index += 1
    depth = 1
    target_tokens: List[str] = []
    while index < len(tokens) and depth > 0:
        token = tokens[index]
        if token == "(":
            depth += 1
            target_tokens.append(token)
        elif token == ")":
            depth -= 1
            if depth > 0:
                target_tokens.append(token)
        else:
            target_tokens.append(token)
        index += 1

    if depth != 0:
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    target_columns = _parse_columns(" ".join(target_tokens))
    invalid_target_columns = [
        col for col in target_columns if not isinstance(col, str) or col == "*"
    ]
    if invalid_target_columns:
        raise ValueError("ON CONFLICT target supports only bare column names")

    if index >= len(tokens) or tokens[index].upper() != "DO":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
    index += 1

    if index >= len(tokens):
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    action = tokens[index].upper()
    index += 1

    if action == "NOTHING":
        if index != len(tokens):
            raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
        if params is not None:
            _bind_params([], params)
        return {
            "target_columns": target_columns,
            "action": "NOTHING",
        }

    if action != "UPDATE":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    if index >= len(tokens) or tokens[index].upper() != "SET":
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
    index += 1

    set_part = " ".join(tokens[index:]).strip()
    if not set_part:
        raise ValueError(f"Invalid ON CONFLICT clause format: {query}")

    assignments = []
    raw_assignments = _split_csv(set_part)
    for assignment in raw_assignments:
        if "=" not in assignment:
            raise ValueError(f"Invalid ON CONFLICT clause format: {query}")
        col, value = assignment.split("=", 1)
        assignments.append(
            {
                "column": col.strip(),
                "value": _parse_upsert_assignment_value(value),
            }
        )

    values_to_bind: List[Any] = []
    for item in assignments:
        assignment_value = item["value"]
        if isinstance(assignment_value, dict):
            values_to_bind.extend(_expression_values_to_bind(assignment_value))
        else:
            values_to_bind.append(assignment_value)

    if params is not None or any(_is_placeholder(value) for value in values_to_bind):
        bound = _bind_params(values_to_bind, params)
        consumed = 0
        for item in assignments:
            assignment_value = item["value"]
            if isinstance(assignment_value, dict):
                consumed += _bind_expression_values(assignment_value, bound, consumed)
            else:
                item["value"] = bound[consumed]
                consumed += 1

    return {
        "target_columns": target_columns,
        "action": "UPDATE",
        "set": assignments,
    }


def _parse_insert(query: str, params: Optional[tuple[Any, ...]]) -> Dict[str, Any]:
    stripped = query.strip()
    upper = stripped.upper()
    insert_prefix = "INSERT INTO "
    if not upper.startswith(insert_prefix):
        raise ValueError(f"Invalid INSERT format: {query}")

    remainder = stripped[len(insert_prefix) :].strip()
    if not remainder:
        raise ValueError(f"Invalid INSERT format: {query}")

    split_index = 0
    if remainder[0] in {'"', "'"}:
        quote_char = remainder[0]
        split_index = 1
        while split_index < len(remainder):
            if remainder[split_index] != quote_char:
                split_index += 1
                continue
            if (
                split_index + 1 < len(remainder)
                and remainder[split_index + 1] == quote_char
            ):
                split_index += 2
                continue
            split_index += 1
            break
    else:
        while (
            split_index < len(remainder)
            and not remainder[split_index].isspace()
            and remainder[split_index] != "("
        ):
            split_index += 1

    table = _parse_table_identifier(remainder[:split_index])
    remainder = remainder[split_index:].strip()
    if not table:
        raise ValueError(f"Invalid INSERT format: {query}")

    columns = None
    if remainder.startswith("("):
        depth = 0
        close_index = -1
        for idx, char in enumerate(remainder):
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    close_index = idx
                    break
            if depth < 0:
                break
        if close_index < 0:
            raise ValueError(f"Invalid INSERT format: {query}")
        cols_part = remainder[1:close_index]
        columns = _parse_columns(cols_part)
        invalid_columns = [
            col for col in columns if not isinstance(col, str) or col == "*"
        ]
        if invalid_columns:
            raise ValueError("INSERT column list supports only bare column names")
        remainder = remainder[close_index + 1 :].strip()

    if not remainder:
        raise ValueError(f"Invalid INSERT format: {query}")

    source_part, on_conflict_clause = _split_insert_on_conflict_clause(remainder)
    if not source_part:
        raise ValueError(f"Invalid INSERT format: {query}")

    remainder_upper = source_part.upper()
    remaining_params: Optional[tuple[Any, ...]] = None
    if remainder_upper.startswith("SELECT"):
        select_params = params
        if params is not None and on_conflict_clause is not None:
            select_param_count = _count_unquoted_placeholders(source_part)
            select_params = params[:select_param_count]
            remaining_params = params[select_param_count:]
        subquery = _parse_select(source_part, select_params)
        if remaining_params is None:
            remaining_params = () if params is not None else None
        values: Any = {"type": "subquery", "query": subquery}
    elif remainder_upper.startswith("VALUES"):
        values_part = source_part[len("VALUES") :].strip()
        if not values_part:
            raise ValueError(f"Invalid INSERT format: {query}")

        raw_rows: List[str] = []
        depth = 0
        row_start = -1
        expect_separator = False
        in_single = False
        in_double = False
        index = 0
        while index < len(values_part):
            char = values_part[index]
            if in_single:
                if char == "'":
                    if index + 1 < len(values_part) and values_part[index + 1] == "'":
                        index += 1
                    else:
                        in_single = False
                index += 1
                continue
            if in_double:
                if char == '"':
                    if index + 1 < len(values_part) and values_part[index + 1] == '"':
                        index += 1
                    else:
                        in_double = False
                index += 1
                continue

            if char == "'":
                in_single = True
                index += 1
                continue
            if char == '"':
                in_double = True
                index += 1
                continue

            if char.isspace() and depth == 0:
                index += 1
                continue

            if depth == 0:
                if expect_separator:
                    if char == ",":
                        expect_separator = False
                        index += 1
                        continue
                    raise ValueError(f"Invalid INSERT format: {query}")
                if char != "(":
                    raise ValueError(f"Invalid INSERT format: {query}")
                row_start = index + 1
                depth = 1
                index += 1
                continue

            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    if row_start < 0:
                        raise ValueError(f"Invalid INSERT format: {query}")
                    raw_rows.append(values_part[row_start:index])
                    expect_separator = True

            if depth < 0:
                raise ValueError(f"Invalid INSERT format: {query}")
            index += 1

        if depth != 0 or in_single or in_double or not raw_rows or not expect_separator:
            raise ValueError(f"Invalid INSERT format: {query}")

        if params is None:
            values = [
                _bind_params(
                    [_parse_value(token) for token in _split_csv(raw_row)], None
                )
                for raw_row in raw_rows
            ]
        else:
            bound_rows: List[List[Any]] = []
            param_index = 0
            for raw_row in raw_rows:
                parsed_row = [_parse_value(token) for token in _split_csv(raw_row)]
                placeholders = [value for value in parsed_row if _is_placeholder(value)]
                needed = len(placeholders)
                row_params = params[param_index : param_index + needed]
                bound_rows.append(_bind_params(parsed_row, row_params))
                param_index += needed
            remaining_params = params[param_index:]
            values = bound_rows
    else:
        raise ValueError(f"Invalid INSERT format: {query}")

    on_conflict = None
    if on_conflict_clause is not None:
        on_conflict = _parse_on_conflict_clause(
            on_conflict_clause, query, remaining_params
        )
        remaining_params = () if params is not None else None

    if (
        params is not None
        and remaining_params is not None
        and len(remaining_params) > 0
    ):
        raise ValueError("Too many parameters for placeholders")

    result: Dict[str, Any] = {
        "action": "INSERT",
        "table": table,
        "columns": columns,
        "values": values,
    }
    if on_conflict is not None:
        result["on_conflict"] = on_conflict
    return result


def _parse_update(query: str, params: Optional[tuple[Any, ...]]) -> Dict[str, Any]:
    upper = query.upper()
    if " SET " not in upper:
        raise ValueError(f"Invalid UPDATE format: {query}")
    set_index = upper.index(" SET ")
    before_set = query[:set_index]
    after_set = query[set_index + len(" SET ") :]
    before_tokens = _tokenize(before_set.strip())
    if len(before_tokens) < 2 or before_tokens[0].upper() != "UPDATE":
        raise ValueError(f"Invalid UPDATE format: {query}")
    table = _parse_table_identifier(before_tokens[1])

    where_part = None
    after_set_tokens = _tokenize(after_set.strip())
    paren_depth = 0
    where_token_index: Optional[int] = None
    for index, token in enumerate(after_set_tokens):
        if token == "(":
            paren_depth += 1
            continue
        if token == ")":
            if paren_depth > 0:
                paren_depth -= 1
            continue
        if paren_depth == 0 and token.upper() == "WHERE":
            where_token_index = index
            break

    if where_token_index is None:
        set_part = " ".join(after_set_tokens)
    else:
        set_part = " ".join(after_set_tokens[:where_token_index])
        where_part = " ".join(after_set_tokens[where_token_index + 1 :])

    assignments = []
    raw_assignments = _split_csv(set_part.strip())
    for assignment in raw_assignments:
        if "=" not in assignment:
            raise ValueError(f"Invalid UPDATE format: {query}")
        col, value = assignment.split("=", 1)
        parsed_value = _parse_assignment_expression(
            value,
            annotate_tables=False,
            error_message=f"Invalid UPDATE format: {query}",
        )
        assignments.append({"column": _parse_column_identifier(col.strip()), "value": parsed_value})

    where = None
    if where_part:
        where = _parse_where_expression(
            where_part,
            params,
            bind_params=False,
            allow_subqueries=True,
            outer_sources={table},
        )

    values_to_bind: List[Any] = []
    for item in assignments:
        assignment_value = item["value"]
        if isinstance(assignment_value, dict):
            values_to_bind.extend(_expression_values_to_bind(assignment_value))
        else:
            values_to_bind.append(assignment_value)
    if where is not None:
        values_to_bind.extend(_where_values_to_bind(where))
    if params is not None or any(_is_placeholder(value) for value in values_to_bind):
        bound = _bind_params(values_to_bind, params)
        consumed = 0
        for item in assignments:
            assignment_value = item["value"]
            if isinstance(assignment_value, dict):
                consumed += _bind_expression_values(assignment_value, bound, consumed)
            else:
                item["value"] = bound[consumed]
                consumed += 1
        if where is not None:
            _bind_where_conditions(where, bound, consumed)

    return {
        "action": "UPDATE",
        "table": table,
        "set": assignments,
        "where": where,
    }


def _parse_delete(query: str, params: Optional[tuple[Any, ...]]) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 3 or tokens[0].upper() != "DELETE" or tokens[1].upper() != "FROM":
        raise ValueError(f"Invalid DELETE format: {query}")
    table = _parse_table_identifier(tokens[2])

    where = None
    if len(tokens) > 3:
        if tokens[3].upper() != "WHERE":
            raise ValueError(f"Invalid DELETE format: {query}")
        where_part = " ".join(tokens[4:])
        where = _parse_where_expression(
            where_part,
            params,
            allow_subqueries=True,
            outer_sources={table},
        )

    return {
        "action": "DELETE",
        "table": table,
        "where": where,
    }
