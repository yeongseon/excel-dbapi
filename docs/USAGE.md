# Usage Guide

## Installation

```bash
pip install excel-dbapi
```

## Basic Example

### Local Excel File

```python
from excel_dbapi.connection import ExcelConnection

# Connect to a local Excel file
with ExcelConnection("example.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    rows = cursor.fetchall()

for row in rows:
    print(row)
```

### Remote Excel File
You can also connect to an Excel file hosted over HTTP/HTTPS:

from excel_dbapi.connection import ExcelConnection

```python
# Connect to a remote Excel file
with ExcelConnection("https://example.com/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    rows = cursor.fetchall()

for row in rows:
    print(row)
```

⚠️ Note
Remote file fetching requires the Excel file to be publicly accessible.
Authentication and private URLs are not supported yet.

## Supported SQL Syntax

- SELECT with column selection and simple WHERE conditions
- INSERT INTO to add new rows
- UPDATE to modify existing rows
- DELETE to remove rows
- CREATE TABLE and DROP TABLE
- Transaction control with commit() and rollback()

See the [Project Roadmap](ROADMAP.md) for upcoming feature support.
