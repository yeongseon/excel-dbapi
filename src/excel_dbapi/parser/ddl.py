from __future__ import annotations

from typing import Any, Dict

from ._constants import _normalize_column_type
from .tokenizer import (
    _parse_table_identifier,
    _split_csv_preserve_empty,
    _tokenize,
)


def _parse_create(query: str) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 3 or tokens[0].upper() != "CREATE" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    table_and_cols = " ".join(tokens[2:]).strip()
    if "(" not in table_and_cols or not table_and_cols.endswith(")"):
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    table_name, cols_part = table_and_cols.split("(", 1)
    table = _parse_table_identifier(table_name)
    if not table:
        raise ValueError("Table name is required in CREATE TABLE")
    cols_part = cols_part.rsplit(")", 1)[0]
    raw_columns = _split_csv_preserve_empty(cols_part)
    empty_indexes = [
        index for index, definition in enumerate(raw_columns) if not definition.strip()
    ]
    has_single_trailing_empty = (
        len(empty_indexes) == 1 and empty_indexes[0] == len(raw_columns) - 1
    )
    if empty_indexes and not has_single_trailing_empty:
        raise ValueError("Malformed column definitions: empty column definition found")
    columns = []
    column_definitions = []
    for col in raw_columns:
        if not col.strip():
            continue
        stripped_col = col.strip()
        parts = stripped_col.split()
        if len(parts) > 2:
            raise ValueError(
                "Malformed column definition: "
                f"{stripped_col!r}. Missing comma between column definitions?"
            )
        column_name = parts[0]
        type_name = "TEXT"
        if len(parts) == 2:
            type_name = _normalize_column_type(parts[1], context="CREATE TABLE")
        columns.append(column_name)
        column_definitions.append({"name": column_name, "type_name": type_name})
    if not columns:
        raise ValueError(f"Invalid CREATE TABLE format: {query}")
    return {
        "action": "CREATE",
        "table": table,
        "columns": columns,
        "column_definitions": column_definitions,
    }


def _parse_drop(query: str) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) != 3 or tokens[0].upper() != "DROP" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid DROP TABLE format: {query}")
    return {
        "action": "DROP",
        "table": _parse_table_identifier(tokens[2]),
    }


def _parse_alter(query: str) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 6 or tokens[0].upper() != "ALTER" or tokens[1].upper() != "TABLE":
        raise ValueError(f"Invalid ALTER TABLE format: {query}")

    table = _parse_table_identifier(tokens[2])
    operation = tokens[3].upper()

    if operation == "ADD":
        if len(tokens) != 7 or tokens[4].upper() != "COLUMN":
            raise ValueError(f"Invalid ALTER TABLE format: {query}")
        type_name = _normalize_column_type(tokens[6], context="ALTER TABLE")
        return {
            "action": "ALTER",
            "table": table,
            "operation": "ADD_COLUMN",
            "column": tokens[5],
            "type_name": type_name,
        }

    if operation == "DROP":
        if len(tokens) != 6 or tokens[4].upper() != "COLUMN":
            raise ValueError(f"Invalid ALTER TABLE format: {query}")
        return {
            "action": "ALTER",
            "table": table,
            "operation": "DROP_COLUMN",
            "column": tokens[5],
        }

    if operation == "RENAME":
        if (
            len(tokens) != 8
            or tokens[4].upper() != "COLUMN"
            or tokens[6].upper() != "TO"
        ):
            raise ValueError(f"Invalid ALTER TABLE format: {query}")
        return {
            "action": "ALTER",
            "table": table,
            "operation": "RENAME_COLUMN",
            "old_column": tokens[5],
            "new_column": tokens[7],
        }

    raise ValueError(f"Invalid ALTER TABLE format: {query}")
