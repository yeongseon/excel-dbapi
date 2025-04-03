# Usage Guide

## Installation

```bash
pip install excel-dbapi
```

## Basic Example

```python
import excel_dbapi

# Connect to an Excel file
conn = excel_dbapi.connect("example.xlsx")
cursor = conn.cursor()

# Read data from a sheet
cursor.execute("SELECT * FROM Sheet1")
rows = cursor.fetchall()

for row in rows:
    print(row)

# Close connection
conn.close()
```

## Supported SQL Syntax

- SELECT with column selection and simple WHERE conditions
- INSERT INTO to add new rows
- UPDATE to modify existing rows
- DELETE to remove rows
- CREATE TABLE and DROP TABLE
- Transaction control with commit() and rollback()

See the [Project Roadmap](ROADMAP.md) for upcoming feature support.
