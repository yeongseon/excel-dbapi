# AGENTS.md — excel-dbapi

> Project knowledge base for AI agents. Last updated: 2026-03-15.

## Project Identity

- **Name**: excel-dbapi
- **Package**: `excel-dbapi` (PyPI), import as `excel_dbapi`
- **License**: MIT (Copyright 2025 Yeongseon Choe)
- **Repository**: https://github.com/yeongseon/excel-dbapi
- **Python**: 3.10+
- **Version**: 1.0.0
- **Stage**: v1.0.0 Stable (feature-complete)

## One-Line Description

PEP 249 (DB-API 2.0) compliant driver that lets you query and manipulate Excel (.xlsx) files using SQL.

## What This Project Does

excel-dbapi treats Excel workbooks as databases and worksheets as tables. It provides a standard Python DB-API 2.0 interface (`connect()`, `cursor()`, `execute()`, `fetchall()`, etc.) so that you can:

1. **Query** Excel data with `SELECT ... FROM SheetName WHERE ...`
2. **Insert** rows with `INSERT INTO SheetName (col1, col2) VALUES (?, ?)`
3. **Update** cells with `UPDATE SheetName SET col = ? WHERE ...`
4. **Delete** rows with `DELETE FROM SheetName WHERE ...`
5. **Create/Drop** sheets with `CREATE TABLE SheetName (col1, col2)` / `DROP TABLE SheetName`
6. **Transaction control** with `commit()` / `rollback()` and snapshot-based state management
7. **Direct workbook access** via the `workbook` property for formatting/styling (openpyxl engine only)

### Relationship with sqlalchemy-excel

excel-dbapi is a **downstream dependency** of [sqlalchemy-excel](https://github.com/yeongseon/sqlalchemy-excel). sqlalchemy-excel uses excel-dbapi as its core Excel I/O layer:

- `ExcelWorkbookSession` wraps an excel-dbapi connection for dual-channel access (SQL + openpyxl workbook)
- `ExcelDbapiReader` uses excel-dbapi cursors for SQL-based data reading
- `ExcelTemplate` and `ExcelExporter` use `ExcelWorkbookSession.open()` for workbook creation

## Tech Stack

| Layer | Technology | Role |
|-------|-----------|------|
| Excel I/O | openpyxl ≥ 3.1 | Read/write xlsx, cell-level access |
| DataFrame I/O | pandas ≥ 2.0 | Alternative DataFrame-based engine |
| SQL Parsing | Custom parser | Parse SQL statements into AST dicts |
| Build | setuptools ≥ 61 | Package build backend |
| Testing | pytest ≥ 8.3 | Unit and integration tests |
| Linting | ruff ≥ 0.11 | Code linting |
| Formatting | black ≥ 25.1 | Code formatting |
| Type checking | mypy ≥ 1.15 | Static type analysis (non-strict mode) |

## Project Structure

```
excel-dbapi/
├── pyproject.toml               # Package config (setuptools), version 1.0.0
├── README.md                    # User-facing documentation
├── AGENTS.md                    # This file
├── PRD.md                       # Product requirements (Korean)
├── ARCH.md                      # Architecture document
├── TDD.md                       # Technical design document
├── CHANGELOG.md                 # Release history
├── LICENSE
├── internal/
│   └── PRD.md                   # Legacy PRD (outdated, kept for reference)
├── docs/
│   ├── INDEX.md                 # Documentation index
│   ├── USAGE.md                 # Usage guide
│   ├── DEVELOPMENT.md           # Development guide
│   └── ROADMAP.md               # Feature roadmap
├── src/
│   └── excel_dbapi/
│       ├── __init__.py           # Module-level DB-API constants, connect()
│       ├── connection.py         # ExcelConnection (PEP 249 Connection)
│       ├── cursor.py             # ExcelCursor (PEP 249 Cursor)
│       ├── exceptions.py         # PEP 249 exception hierarchy
│       └── engine/
│           ├── base.py           # BaseEngine ABC
│           ├── openpyxl_engine.py # OpenpyxlEngine (default)
│           ├── pandas_engine.py  # PandasEngine (alternative)
│           ├── parser.py         # SQL parser (SELECT/INSERT/UPDATE/DELETE/CREATE/DROP)
│           ├── executor.py       # Query dispatcher (case-insensitive table lookup)
│           ├── openpyxl_executor.py  # Executor for openpyxl engine
│           ├── pandas_executor.py    # Executor for pandas engine
│           └── result.py         # ExecutionResult dataclass
├── tests/                        # 85 tests
│   ├── test_connection.py
│   ├── test_cursor.py
│   ├── test_engine.py
│   ├── test_parser.py
│   └── ...
└── examples/
    ├── basic_usage.py
    ├── write_operations.py
    ├── transactions.py
    ├── advanced_query.py
    └── pandas_engine.py
```

## DB-API 2.0 Module-Level Constants

```python
import excel_dbapi

excel_dbapi.apilevel     # "2.0"
excel_dbapi.threadsafety  # 1 (threads may share module, not connections)
excel_dbapi.paramstyle    # "qmark" (question mark placeholders: WHERE id = ?)
```

## Key Design Decisions

1. **PEP 249 compliance** — Full DB-API 2.0 interface: `connect()`, `Connection`, `Cursor`, standard exceptions
2. **Dual engine architecture** — OpenpyxlEngine (default, cell-level) and PandasEngine (DataFrame-based), selectable at connect time
3. **Custom SQL parser** — Lightweight recursive-descent parser; no external SQL parsing dependency
4. **Case-insensitive table lookup** — Sheet names are matched case-insensitively via `table.lower()` mapping
5. **Unquoted table names** — SQL uses `SELECT * FROM Sheet1`, NOT `"Sheet1"` — quotes cause lookup failure
6. **Snapshot-based transactions** — `commit()` persists to disk; `rollback()` restores from in-memory snapshot
7. **Atomic file save** — OpenpyxlEngine saves via `tempfile.NamedTemporaryFile` + `os.replace` to prevent corruption
8. **`create` flag** — `connect(path, create=True)` creates the file if it doesn't exist
9. **`data_only` flag** — Controls whether openpyxl reads cached formula values (`True`, default) or formulas (`False`)
10. **`workbook` property** — OpenpyxlEngine exposes the raw openpyxl `Workbook` object; PandasEngine raises `NotSupportedError`
11. **`check_closed` decorator** — Both Connection and Cursor use a decorator to guard against operations on closed objects
12. **Autocommit default** — `autocommit=True` by default; write operations auto-save. Rollback disabled in autocommit mode.

## Exception Hierarchy (PEP 249)

```python
Exception
├── Warning                     # Important warnings (e.g., data truncations)
└── Error                       # Base for all DB-API errors
    ├── InterfaceError           # Connection/cursor interface problems
    └── DatabaseError            # Database-related errors
        ├── DataError            # Value out of range, type mismatch
        ├── OperationalError     # Connection issues, file I/O problems
        ├── IntegrityError       # Relational integrity violations
        ├── InternalError        # Internal engine errors
        ├── ProgrammingError     # Bad SQL syntax, unknown columns
        └── NotSupportedError    # Unsupported operations (e.g., rollback in autocommit)
```

## Public API

### `connect()` — Module-level constructor

```python
from excel_dbapi import connect

conn = connect(
    file_path="data.xlsx",      # Path to Excel file
    engine="openpyxl",          # "openpyxl" (default) or "pandas"
    autocommit=True,            # Auto-save on write operations (default: True)
    create=False,               # Create file if missing (default: False)
    data_only=True,             # Read cached formula values (default: True)
)
```

### `ExcelConnection` — PEP 249 Connection

```python
from excel_dbapi.connection import ExcelConnection

conn = ExcelConnection("data.xlsx")
cursor = conn.cursor()           # Create a new cursor
conn.commit()                    # Persist changes and take new snapshot
conn.rollback()                  # Restore from snapshot (autocommit=False only)
conn.close()                     # Close the connection
conn.closed                      # bool — whether connection is closed
conn.engine_name                 # str — "OpenpyxlEngine" or "PandasEngine"
conn.workbook                    # openpyxl Workbook (openpyxl engine only)

# Context manager support
with ExcelConnection("data.xlsx") as conn:
    ...
```

### `ExcelCursor` — PEP 249 Cursor

```python
cursor = conn.cursor()
cursor.execute(query, params=None)          # Execute a single query
cursor.executemany(query, seq_of_params)    # Execute query for each param tuple
cursor.fetchone()                           # Fetch one row as tuple, or None
cursor.fetchall()                           # Fetch all remaining rows
cursor.fetchmany(size=None)                 # Fetch `size` rows (default: arraysize)
cursor.description                          # Column metadata (7-tuples per PEP 249)
cursor.rowcount                             # Number of affected/returned rows
cursor.lastrowid                            # Last inserted row index
cursor.arraysize                            # Default fetch size (default: 1)
cursor.close()                              # Close the cursor
```

## SQL Support

### Supported Statements

| Statement | Syntax | Notes |
|-----------|--------|-------|
| SELECT | `SELECT col1, col2 FROM Sheet1 WHERE ... ORDER BY col ASC/DESC LIMIT N` | `*` for all columns |
| INSERT | `INSERT INTO Sheet1 (col1, col2) VALUES (?, ?)` | Column list optional |
| UPDATE | `UPDATE Sheet1 SET col1 = ? WHERE ...` | Multiple SET assignments |
| DELETE | `DELETE FROM Sheet1 WHERE ...` | WHERE optional (deletes all) |
| CREATE TABLE | `CREATE TABLE SheetName (col1, col2)` | Creates new worksheet |
| DROP TABLE | `DROP TABLE SheetName` | Removes worksheet |

### WHERE Clause

- Comparison operators: `=`, `==`, `!=`, `<>`, `>`, `>=`, `<`, `<=`
- Logical operators: `AND`, `OR`
- Parameter binding: `?` placeholders with `qmark` paramstyle
- Type coercion: Numeric strings compared numerically when both sides can be parsed as numbers

### Parameter Binding

```python
# Positional parameters with ?
cursor.execute("SELECT * FROM Sheet1 WHERE id = ?", (42,))
cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)", (1, "Alice"))
cursor.execute("UPDATE Sheet1 SET name = ? WHERE id = ?", ("Bob", 1))
```

## Coding Conventions

- **Type hints**: Public functions typed. Some internal functions use `Any` due to openpyxl/pandas interop.
- **Docstrings**: Google style for public classes/methods.
- **Build**: setuptools (not hatch/poetry).
- **Testing**: pytest with fixtures. 85 tests covering both engines.
- **Linting**: ruff for lint, black for format.
- **Type checking**: mypy with `strict = false`, `ignore_missing_imports = true`.
- **Error handling**: All exceptions follow PEP 249 hierarchy. Never bare `except:`.
- **Imports**: No `from __future__ import annotations` requirement (unlike sqlalchemy-excel).

## Testing Strategy

- **Unit tests**: Connection, Cursor, Parser, each Executor independently
- **Engine parity tests**: Same operations verified on both OpenpyxlEngine and PandasEngine
- **Transaction tests**: Commit/rollback/snapshot lifecycle
- **SQL syntax tests**: Valid and invalid SQL parsing edge cases
- **Error handling**: Closed connection/cursor, missing sheets, bad SQL
- **Total**: 85 tests, all passing
- **CI**: GitHub Actions (matrix testing planned)

## Environment Setup

```bash
# Development
pip install -e ".[dev]"

# Run tests
pytest

# Or with Makefile
make setup    # Create virtualenv + install
make test     # Run tests with coverage

# Lint + format
ruff check src/ tests/
black src/ tests/

# Type check
mypy src/
```

## Important Notes for AI Agents

1. **Package name** is `excel-dbapi` (hyphen), import name is `excel_dbapi` (underscore)
2. **Table names in SQL must be UNQUOTED**: `SELECT * FROM Sheet1`, NOT `SELECT * FROM "Sheet1"` — quotes cause lookup failure
3. **Sheet name matching is case-insensitive**: `FROM sheet1` matches a sheet named `Sheet1`
4. **OpenpyxlEngine** has a `workbook` property (returns openpyxl `Workbook`), but **PandasEngine does NOT** — raises `NotSupportedError`
5. **Autocommit is ON by default**. Write operations (INSERT/UPDATE/DELETE/CREATE/DROP) auto-save to disk. Set `autocommit=False` to enable manual commit/rollback.
6. **Rollback is disabled in autocommit mode** — calling `rollback()` with `autocommit=True` raises `NotSupportedError`
7. **Parser is custom** — no external SQL library dependency. Supports a subset of SQL, not full SQL92.
8. **Both engines require openpyxl**: PandasEngine uses pandas for in-memory data but openpyxl for initial file creation and reads.
9. **File save is atomic** (OpenpyxlEngine): writes to temp file, then `os.replace()` — safe against crashes.
10. **`executemany()` with `autocommit=False`** uses snapshot for atomic batch: if any param fails, all are rolled back.
11. **`pyproject.toml` version is 1.0.0**, but PyPI has a v2.0.0 published in error (withdrawn).
12. **Build system is setuptools**, not hatchling (unlike sqlalchemy-excel which uses hatchling).
13. **mypy is non-strict** (`strict = false`), unlike sqlalchemy-excel which uses `strict = true`.
14. **`description` format** follows PEP 249: sequence of 7-tuples `(name, type_code, display_size, internal_size, precision, scale, null_ok)`.
15. **ORDER BY with unknown column** raises `ValueError`. Direction must be `ASC` or `DESC`.
16. **LIMIT** must be a positive integer. Non-integer LIMIT raises `ValueError`.
17. **INSERT into sheet without headers** raises `ValueError` — sheets must have a header row.
18. **CREATE TABLE on existing sheet** raises `ValueError`. **DROP TABLE on missing sheet** raises `ValueError`.
