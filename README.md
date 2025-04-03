
# excel-dbapi

![CI](https://github.com/your-username/excel-dbapi/actions/workflows/ci.yml/badge.svg)
[![codecov](https://codecov.io/gh/your-username/excel-dbapi/branch/main/graph/badge.svg)](https://codecov.io/gh/your-username/excel-dbapi)

A lightweight, Python DB-API 2.0 compliant connector for Excel files.

---

## Features

- Python DB-API 2.0 compliant interface
- Query Excel files using SQL syntax
- Supports SELECT (INSERT, UPDATE, DELETE planned)
- Sheet-to-Table mapping
- Pandas & Openpyxl engine selector
- Transaction simulation (planned)
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

# Using pandas engine
with ExcelConnection("path/to/sample.xlsx", engine="pandas") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())
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

## Planned Features

- Write operations (INSERT, UPDATE, DELETE)
- DDL support (CREATE TABLE, DROP TABLE)
- Transaction simulation
- Advanced SQL condition support (WHERE, ORDER BY, LIMIT)
- Remote file connection support
- SQLAlchemy Dialect

See [Project Roadmap](docs/ROADMAP.md) for details.

---

## Documentation

- [Usage Guide](docs/USAGE.md)
- [Development Guide](docs/DEVELOPMENT.md)
- [Project Roadmap](docs/ROADMAP.md)

---

## License

MIT License
