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

## Write Operations

```python
with ExcelConnection("tests/data/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (1, 'Alice')")
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    cursor.execute("DELETE FROM Sheet1 WHERE id = 2")
```

## Bulk Inserts

```python
with ExcelConnection("tests/data/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.executemany(
        "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
        [(2, "Bob"), (3, "Cara")],
    )
```

## DDL

```python
with ExcelConnection("tests/data/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE NewSheet (id, name)")
    cursor.execute("DROP TABLE NewSheet")
```

## Transactions

Autocommit is enabled by default. To use transactions:

```python
with ExcelConnection("tests/data/sample.xlsx", autocommit=False) as conn:
    cursor = conn.cursor()
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    conn.rollback()
```

When autocommit is enabled, `rollback()` is not supported.

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

## Limitations

- `PandasEngine` rewrites workbooks and may drop formatting, charts, and formulas.
- `OpenpyxlEngine` loads with `data_only=True`, so formulas are evaluated to values when reading.

## Cursor Metadata

```python
with ExcelConnection("tests/data/sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM Sheet1")
    print(cursor.description)
    print(cursor.rowcount)
```
