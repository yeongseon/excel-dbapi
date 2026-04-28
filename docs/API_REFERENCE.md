# API Reference

This document is generated from the current `src/excel_dbapi` public API and
module docstrings. It focuses on the stable user-facing surface.

## Top-level module: `excel_dbapi`

### `connect(file_path, engine=None, autocommit=True, create=False, backup=False, backup_dir=None, data_only=True, sanitize_formulas=True, credential=None, warn_rows=None, **backend_options) -> ExcelConnection`

Create a DB-API connection to a local workbook (`.xlsx`) or a DSN.

- `file_path`: local path or DSN (for example `msgraph://drives/{id}/items/{id}`)
- `engine`: `openpyxl`, `pandas`, `graph`, or `None` for DSN auto-detection
- `autocommit`: save write operations automatically when `True`
- `create`: create workbook if missing (backend-dependent)
- `backup`: if `True`, create a timestamped backup before the first mutating operation (local files only)
- `backup_dir`: custom backup directory (default: `.excel-dbapi-backups/` next to the workbook)
- `data_only`: read formula cached values instead of formula text
- `sanitize_formulas`: escape formula-like user input on write
- `credential`: optional credential/token provider for cloud backends
- `warn_rows`: emit a `UserWarning` when a sheet exceeds this row count (default: disabled)
- `backend_options`: additional backend-specific options

Example:

```python
import excel_dbapi

conn = excel_dbapi.connect("sample.xlsx", engine="openpyxl")
cur = conn.cursor()
cur.execute("SELECT * FROM Sheet1")
rows = cur.fetchall()
conn.close()
```

## `ExcelConnection`

PEP 249-style connection object (`excel_dbapi.connection.ExcelConnection`).

### Constructor

`ExcelConnection(file_path, engine=None, autocommit=True, create=False, backup=False, backup_dir=None, data_only=True, sanitize_formulas=True, credential=None, warn_rows=None, **backend_options)`

### Methods

- `cursor() -> ExcelCursor`
- `execute(query: str, params: Sequence[Any] | None = None) -> ExecutionResult`
- `commit() -> None`
- `rollback() -> None`
- `close() -> None`

### Properties

- `engine_name: str`
- `workbook: Any` (backend-dependent; may raise `NotSupportedError`)
- `closed: bool`

Example with context manager:

```python
from excel_dbapi import connect

with connect("sample.xlsx", autocommit=False) as conn:
    cur = conn.cursor()
    cur.execute("UPDATE users SET name = 'Ann' WHERE id = 1")
    conn.commit()
```

## `ExcelCursor`

PEP 249-style cursor object (`excel_dbapi.cursor.ExcelCursor`).

### Methods

- `execute(query: str, params: Sequence[Any] | None = None) -> ExcelCursor`
- `executemany(query: str, seq_of_params: Iterable[Sequence[Any]]) -> ExcelCursor`
- `fetchone() -> tuple[Any, ...] | None`
- `fetchall() -> list[tuple[Any, ...]]`
- `fetchmany(size: int | None = None) -> list[tuple[Any, ...]]`
- `close() -> None`
- `setinputsizes(...) -> None` (compatibility no-op)
- `setoutputsize(...) -> None` (compatibility no-op)

### Attributes

- `description`
- `rowcount`
- `lastrowid`
- `arraysize`
- `closed`

Example:

```python
from excel_dbapi import connect

with connect("sample.xlsx") as conn:
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM users WHERE id > ?", (10,))
    for row in cur.fetchall():
        print(row)
```

## Exceptions

Defined in `excel_dbapi.exceptions`:

- `Error`
- `Warning`
- `InterfaceError`
- `DatabaseError`
- `DataError`
- `OperationalError`
- `IntegrityError`
- `InternalError`
- `ProgrammingError`
- `NotSupportedError`

Example:

```python
from excel_dbapi import connect
from excel_dbapi.exceptions import ProgrammingError

try:
    with connect("sample.xlsx") as conn:
        conn.cursor().execute("SELECT missing_column FROM users")
except ProgrammingError as exc:
    print(f"Invalid query: {exc}")
```

## Type System and Result Shapes

### Inferred logical types

`excel_dbapi.reflection.get_columns()` infers:

- `TEXT`
- `BOOLEAN`
- `INTEGER`
- `FLOAT`
- `DATE`
- `DATETIME`

These are returned as metadata dictionaries:

```python
[
    {"name": "id", "type": "INTEGER", "nullable": False},
    {"name": "created_at", "type": "DATETIME", "nullable": True},
]
```

### Backend table container

`excel_dbapi.engines.base.TableData`:

- `headers: list[str]`
- `rows: list[list[Any]]`

### Execution result container

`excel_dbapi.engines.result.ExecutionResult`:

- `action: str`
- `rows: list[tuple[Any, ...]]`
- `description: Description`
- `rowcount: int`
- `lastrowid: int | None`

`Description` follows DB-API cursor description tuple shape.

## Reflection helpers

Module: `excel_dbapi.reflection`

- `list_tables(connection, include_meta=False) -> list[str]`
- `has_table(connection, table_name) -> bool`
- `get_columns(connection, table_name, sample_size=100) -> list[dict[str, Any]]`
- `read_table_metadata(connection, table_name) -> list[dict[str, Any]] | None`
- `write_table_metadata(connection, table_name, columns) -> None`
- `remove_table_metadata(connection, table_name) -> None`
- `METADATA_SHEET = "__excel_meta__"`
