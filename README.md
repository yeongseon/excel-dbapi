# excel-dbapi

![CI](https://github.com/your-username/excel-dbapi/actions/workflows/ci.yml/badge.svg)
[![codecov](https://codecov.io/gh/your-username/excel-dbapi/branch/main/graph/badge.svg)](https://codecov.io/gh/your-username/excel-dbapi)

A lightweight, Python DB-API 2.0 compliant connector for Excel files.

---

## Features

- Python DB-API 2.0 compliant interface
- Query Excel files using SQL syntax
- Supports SELECT, INSERT, UPDATE, DELETE
- Transaction support
- SQLAlchemy Dialect integration (upcoming)

---

## Installation

```bash
pip install excel-dbapi
```

See [CHANGELOG](CHANGELOG.md) for release history.

---

## Quick Start

### Local file connection

```python
from excel_dbapi.connection import ExcelConnection

with ExcelConnection("path/to/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())
```

### Remote file connection

```python
from excel_dbapi.connection import ExcelConnection

with ExcelConnection("https://example.com/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())
```

⚠️ Note
Remote file fetching requires the Excel file to be publicly accessible.
Authentication and private URLs are not supported yet.

## Documentation

- [Usage Guide](docs/USAGE.md)
- [Development Guide](docs/DEVELOPMENT.md)
- [Project Roadmap](docs/ROADMAP.md)

---

## License

MIT License
