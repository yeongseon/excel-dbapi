
# excel-dbapi

![CI](https://github.com/your-username/excel-dbapi/actions/workflows/ci.yml/badge.svg)
[![codecov](https://codecov.io/gh/your-username/excel-dbapi/branch/main/graph/badge.svg)](https://codecov.io/gh/your-username/excel-dbapi)

A lightweight, Python DB-API 2.0 compliant connector for Excel files.

---

## Features

- Python DB-API 2.0 compliant interface
- Query Excel files using SQL syntax
- Supports SELECT, INSERT, UPDATE, DELETE
- Basic DDL support (CREATE TABLE, DROP TABLE)
- WHERE conditions with AND/OR and comparison operators
- ORDER BY and LIMIT for SELECT
- Sheet-to-Table mapping
- Pandas & Openpyxl engine selector
- Transaction simulation (commit/rollback)
- SQLAlchemy Dialect integration (planned)

---

## Installation

```bash
pip install excel-dbapi
```

See [CHANGELOG](CHANGELOG.md) for release history.

---

## Quick Start

### Basic Usage (Local File)

```python
from excel_dbapi.connection import ExcelConnection

# Using default engine (openpyxl)
with ExcelConnection("path/to/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())

# Insert a row
with ExcelConnection("path/to/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (1, 'Alice')")

# Using pandas engine
with ExcelConnection("path/to/sample.xlsx", engine="pandas") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())

# Update and delete rows
with ExcelConnection("path/to/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    cursor.execute("DELETE FROM Sheet1 WHERE id = 2")

# Create and drop sheets
with ExcelConnection("path/to/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE NewSheet (id, name)")
    cursor.execute("DROP TABLE NewSheet")
```

### Engine Options

| Engine    | Description                  | Dependency  |
|---------|------------------------------|--------------|
| openpyxl (default) | Fast sheet access (read-only) | openpyxl   |
| pandas  | DataFrame based operations   | pandas, openpyxl |

You can explicitly specify the engine using:

```python
conn = ExcelConnection("sample.xlsx", engine="openpyxl")
conn = ExcelConnection("sample.xlsx", engine="pandas")
```

---

## Transaction Example

```python
with ExcelConnection("path/to/sample.xlsx", autocommit=False) as conn:
    cursor = conn.cursor()
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    conn.rollback()
```

## Cursor Metadata

```python
with ExcelConnection("path/to/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM Sheet1")
    print(cursor.description)
    print(cursor.rowcount)
```

## Planned Features

- Remote file connection support
- SQLAlchemy Dialect

See [Project Roadmap](docs/ROADMAP.md) for details.

---

## Documentation

- [Usage Guide](docs/USAGE.md)
- [Development Guide](docs/DEVELOPMENT.md)
- [Project Roadmap](docs/ROADMAP.md)

## Examples

- `examples/basic_usage.py`
- `examples/write_operations.py`
- `examples/transactions.py`
- `examples/advanced_query.py`
- `examples/pandas_engine.py`

---

## License

MIT License
