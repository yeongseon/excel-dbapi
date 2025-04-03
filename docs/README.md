# Usage Guide

## Installation

```bash
pip install excel-dbapi
```

## Basic Example

```python
import excel_dbapi

conn = excel_dbapi.connect("sample.xlsx")
cursor = conn.cursor()
cursor.execute("SELECT * FROM Sheet1")
print(cursor.fetchall())
conn.close()
```

## Supported SQL Syntax

- SELECT
- INSERT
- UPDATE
- DELETE
- CREATE TABLE, DROP TABLE
- Transaction support

For upcoming features, see [Project Roadmap](ROADMAP.md).
