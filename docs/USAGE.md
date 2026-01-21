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

- SELECT with WHERE (AND/OR, comparison operators), ORDER BY, LIMIT
- INSERT, UPDATE, DELETE
- CREATE TABLE, DROP TABLE

## Transactions

Autocommit is enabled by default. To use transactions:

```python
with ExcelConnection("tests/data/sample.xlsx", autocommit=False) as conn:
    cursor = conn.cursor()
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    conn.rollback()
```

For upcoming features, see [Project Roadmap](ROADMAP.md).

## Advanced Examples

```python
with ExcelConnection("tests/data/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, name FROM Sheet1 WHERE id >= ? ORDER BY id DESC LIMIT 1",
        (1,),
    )
    print(cursor.fetchall())
```
