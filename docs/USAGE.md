# Usage Guide

## Installation

```bash
pip install excel-dbapi
```

## Basic Example

```python
from excel_dbapi.connection import ExcelConnection

# Open a connection to the Excel file
with ExcelConnection("tests/data/sample.xlsx") as conn:
    # Create a cursor
    cursor = conn.cursor()
    
    # Execute a query
    cursor.execute("SELECT * FROM Sheet1")
    
    # Fetch results
    results = cursor.fetchall()
    
    # Display results
    for row in results:
        print(row)

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
