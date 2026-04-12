"""Reflection helpers for dialect integration."""

from __future__ import annotations

import datetime
from typing import Any

from excel_dbapi.engines.base import TableData

METADATA_SHEET = "__excel_meta__"


def list_tables(connection: Any, include_meta: bool = False) -> list[str]:
    """Return worksheet names, excluding metadata sheet by default."""
    sheets = connection.engine.list_sheets()
    if not include_meta:
        sheets = [sheet for sheet in sheets if sheet != METADATA_SHEET]
    return sheets


def has_table(connection: Any, table_name: str) -> bool:
    """Check if a worksheet exists (case-insensitive)."""
    return table_name.lower() in {
        sheet.lower() for sheet in connection.engine.list_sheets()
    }


def get_columns(
    connection: Any, table_name: str, sample_size: int = 100
) -> list[dict[str, Any]]:
    """Return column metadata by sampling data rows."""
    data = connection.engine.read_sheet(table_name)
    columns: list[dict[str, Any]] = []
    for index, header in enumerate(data.headers):
        col_values = [row[index] for row in data.rows[:sample_size] if index < len(row)]
        inferred = _infer_type(col_values)
        columns.append(
            {
                "name": header,
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
    types = set()
    for value in non_null:
        if isinstance(value, bool):
            types.add("BOOLEAN")
        elif isinstance(value, int):
            types.add("INTEGER")
        elif isinstance(value, float):
            types.add("FLOAT")
        elif isinstance(value, datetime.datetime):
            types.add("DATETIME")
        elif isinstance(value, datetime.date):
            types.add("DATE")
        else:
            types.add("TEXT")

    if types == {"INTEGER"}:
        return {"type": "INTEGER", "nullable": nullable}
    if types <= {"INTEGER", "FLOAT"}:
        return {"type": "FLOAT", "nullable": nullable}
    if types == {"BOOLEAN"}:
        return {"type": "BOOLEAN", "nullable": nullable}
    if types == {"DATE"}:
        return {"type": "DATE", "nullable": nullable}
    if types == {"DATETIME"} or types == {"DATE", "DATETIME"}:
        return {"type": "DATETIME", "nullable": nullable}
    return {"type": "TEXT", "nullable": nullable}


def write_table_metadata(
    connection: Any, table_name: str, columns: list[dict[str, Any]]
) -> None:
    """Write column metadata to the hidden metadata sheet."""
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
    new_rows = [row for row in existing.rows if row[0] != table_name]

    for index, column in enumerate(columns):
        new_rows.append(
            [
                table_name,
                column["name"],
                index + 1,
                column.get("type_name", "TEXT"),
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
    connection: Any, table_name: str
) -> list[dict[str, Any]] | None:
    """Read column metadata from the metadata sheet."""
    sheets = connection.engine.list_sheets()
    if METADATA_SHEET not in sheets:
        return None

    data = connection.engine.read_sheet(METADATA_SHEET)
    entries = [row for row in data.rows if row[0] == table_name]
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


def remove_table_metadata(connection: Any, table_name: str) -> None:
    """Remove metadata for a table from the metadata sheet."""
    sheets = connection.engine.list_sheets()
    if METADATA_SHEET not in sheets:
        return

    data = connection.engine.read_sheet(METADATA_SHEET)
    new_rows = [row for row in data.rows if row[0] != table_name]

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
