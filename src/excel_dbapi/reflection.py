"""Reflection helpers for dialect integration."""

from __future__ import annotations

import datetime
from collections import Counter
from typing import TYPE_CHECKING, Any

from excel_dbapi.engines.base import TableData

if TYPE_CHECKING:
    from excel_dbapi.connection import ExcelConnection

METADATA_SHEET = "__excel_meta__"


def list_tables(connection: ExcelConnection, include_meta: bool = False) -> list[str]:
    """Return worksheet names, excluding metadata sheet by default."""
    sheets: list[str] = connection.engine.list_sheets()
    if not include_meta:
        sheets = [sheet for sheet in sheets if sheet != METADATA_SHEET]
    return sheets


def has_table(connection: ExcelConnection, table_name: str) -> bool:
    """Check if a worksheet exists (case-insensitive)."""
    return _resolve_sheet_name(connection, table_name) is not None


def _resolve_sheet_name(connection: ExcelConnection, table_name: str) -> str | None:
    for sheet_name in connection.engine.list_sheets():
        if sheet_name.lower() == table_name.lower():
            return sheet_name
    return None


def get_columns(
    connection: ExcelConnection, table_name: str, sample_size: int | None = 100
) -> list[dict[str, Any]]:
    """Return column metadata by sampling data rows."""
    resolved_table_name = _resolve_sheet_name(connection, table_name)
    if resolved_table_name is None:
        raise ValueError(f"Sheet '{table_name}' not found in Excel")

    data = connection.engine.read_sheet(resolved_table_name)
    columns: list[dict[str, Any]] = []
    sampled_rows = data.rows if sample_size is None else data.rows[:sample_size]
    for index, header in enumerate(data.headers):
        col_values = [row[index] for row in sampled_rows if index < len(row)]
        inferred = _infer_type(col_values)
        columns.append(
            {
                "name": header,
                "type_name": inferred["type"],
                "type": inferred["type"],
                "nullable": inferred["nullable"],
            }
        )
    return columns


def _infer_type(values: list[Any]) -> dict[str, Any]:
    """Infer column type from sample values."""
    non_null = [value for value in values if value is not None]
    if not non_null:
        return {"type": "TEXT", "nullable": True}

    nullable = len(non_null) < len(values)
    type_names = [_classify_value_type(value) for value in non_null]
    counts = Counter(type_names)
    unique_types = set(counts)

    if len(unique_types) == 1:
        return {"type": type_names[0], "nullable": nullable}

    if unique_types <= {"INTEGER", "FLOAT"}:
        return {"type": "FLOAT", "nullable": nullable}

    if unique_types <= {"DATE", "DATETIME"}:
        return {"type": "DATETIME", "nullable": nullable}

    dominant_type, dominant_count = counts.most_common(1)[0]
    if dominant_count / len(non_null) > 0.8:
        return {"type": dominant_type, "nullable": nullable}

    return {"type": "TEXT", "nullable": nullable}


def _classify_value_type(value: Any) -> str:
    if isinstance(value, bool):
        return "BOOLEAN"
    if isinstance(value, int):
        return "INTEGER"
    if isinstance(value, float):
        return "FLOAT"
    if isinstance(value, datetime.datetime):
        return "DATETIME"
    if isinstance(value, datetime.date):
        return "DATE"
    return "TEXT"


def write_table_metadata(
    connection: ExcelConnection, table_name: str, columns: list[dict[str, Any]]
) -> None:
    """Write column metadata to the internal metadata sheet."""
    engine = connection.engine
    sheets = engine.list_sheets()

    if METADATA_SHEET not in sheets:
        engine.create_sheet(
            METADATA_SHEET,
            [
                "table_name",
                "column_name",
                "ordinal",
                "type_name",
                "nullable",
                "primary_key",
            ],
        )

    existing = engine.read_sheet(METADATA_SHEET)
    table_name_key = table_name.casefold()
    new_rows = [
        row
        for row in existing.rows
        if not row or str(row[0]).casefold() != table_name_key
    ]

    for index, column in enumerate(columns):
        type_name = column.get("type_name", column.get("type", "TEXT"))
        new_rows.append(
            [
                table_name,
                column["name"],
                index + 1,
                type_name,
                str(column.get("nullable", True)),
                str(column.get("primary_key", False)),
            ]
        )

    engine.write_sheet(
        METADATA_SHEET,
        TableData(
            headers=[
                "table_name",
                "column_name",
                "ordinal",
                "type_name",
                "nullable",
                "primary_key",
            ],
            rows=new_rows,
        ),
    )


def read_table_metadata(
    connection: ExcelConnection, table_name: str
) -> list[dict[str, Any]] | None:
    """Read column metadata from the metadata sheet."""
    sheets = connection.engine.list_sheets()
    if METADATA_SHEET not in sheets:
        return None

    data = connection.engine.read_sheet(METADATA_SHEET)
    table_name_key = table_name.casefold()
    entries = [
        row
        for row in data.rows
        if row and str(row[0]).casefold() == table_name_key
    ]
    if not entries:
        return None

    entries.sort(key=lambda row: int(row[2]) if row[2] is not None else 0)

    return [
        {
            "name": row[1],
            "type_name": row[3],
            "nullable": row[4] == "True",
            "primary_key": row[5] == "True",
        }
        for row in entries
    ]


def remove_table_metadata(connection: ExcelConnection, table_name: str) -> None:
    """Remove metadata for a table from the metadata sheet."""
    sheets = connection.engine.list_sheets()
    if METADATA_SHEET not in sheets:
        return

    data = connection.engine.read_sheet(METADATA_SHEET)
    table_name_key = table_name.casefold()
    new_rows = [
        row
        for row in data.rows
        if not row or str(row[0]).casefold() != table_name_key
    ]

    connection.engine.write_sheet(
        METADATA_SHEET,
        TableData(
            headers=[
                "table_name",
                "column_name",
                "ordinal",
                "type_name",
                "nullable",
                "primary_key",
            ],
            rows=new_rows,
        ),
    )
