# Usage Guide

## Installation

```bash
pip install excel-dbapi
```

## Basic Example

```python
import excel_dbapi

# Connect to Excel file with default engine (openpyxl)
conn = excel_dbapi.connect("sample.xlsx")
cursor = conn.cursor()
cursor.execute("SELECT * FROM [Sheet1$]")
print(cursor.fetchall())
conn.close()

# You can also use pandas engine
conn = excel_dbapi.connect("sample.xlsx", engine="pandas")
cursor = conn.cursor()
cursor.execute("SELECT * FROM [Sheet1$]")
print(cursor.fetchall())
conn.close()

```

## Supported SQL Syntax

- SELECT (Basic query with optional WHERE clause)
Currently, only SELECT queries are supported.

The following features are planned for future releases:
- INSERT
- UPDATE
- DELETE
- CREATE TABLE, DROP TABLE
- Transaction support

For upcoming features, see [Project Roadmap](ROADMAP.md).
