# Usage Guide

## Installation

```bash
pip install excel-dbapi
```

## Basic Example

```python
from excel_dbapi.connection import ExcelConnection

# Open a connection to the Excel file
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()

    # Execute a query
    cursor.execute("SELECT * FROM Sheet1")

    # Fetch results
    results = cursor.fetchall()

    for row in results:
        print(row)
```

## Supported SQL Syntax

excel-dbapi implements a deliberate subset of SQL. For the complete formal grammar, see
[SQL_SPEC.md](SQL_SPEC.md).

- **SELECT** with WHERE (AND/OR, comparison operators), ORDER BY, LIMIT
- **INSERT INTO** with VALUES
- **UPDATE** with SET and WHERE
- **DELETE FROM** with WHERE
- **CREATE TABLE**, **DROP TABLE**

### Excluded by Design

JOIN, GROUP BY, HAVING, DISTINCT, OFFSET, subqueries, CTEs, and aggregate functions
are not supported. See [Limitations](#limitations) below.

## Write Operations

```python
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (1, 'Alice')")
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    cursor.execute("DELETE FROM Sheet1 WHERE id = 2")
```

## Bulk Inserts

```python
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.executemany(
        "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
        [(2, "Bob"), (3, "Cara")],
    )
```

## DDL

```python
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE NewSheet (id, name)")
    cursor.execute("DROP TABLE NewSheet")
```

## Transactions

Autocommit is enabled by default. To use manual transactions:

```python
with ExcelConnection("sample.xlsx", autocommit=False) as conn:
    cursor = conn.cursor()
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    conn.rollback()  # restores the in-memory snapshot (reverts uncommitted changes)
```

When `autocommit=True` (default), changes are written to disk immediately and
`rollback()` raises `NotSupportedError`.

When `autocommit=False`, `rollback()` restores the in-memory snapshot taken at
connection open (or the last `commit()`). This is **not** a WAL — it restores an
in-memory copy, not a durable transaction log.

## Advanced Examples

```python
with ExcelConnection("sample.xlsx") as conn:
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
- SQL subset deliberately excludes JOIN, GROUP BY, HAVING, DISTINCT, and subqueries.
- No concurrent write support — use a single-writer model.
- Rollback restores an in-memory snapshot, not a durable transaction log.

## Security and Parameter Binding

Always prefer placeholders (`?`) for dynamic values:

```python
cursor.execute("SELECT * FROM Sheet1 WHERE id = ?", (1,))
```

Avoid string interpolation for SQL parameters. excel-dbapi uses **qmark paramstyle** (`?`).

## Cursor Metadata

```python
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM Sheet1")
    print(cursor.description)
    print(cursor.rowcount)
```

## Further Reading

- [SQL Specification (EBNF grammar)](SQL_SPEC.md)
- [10-Minute Quickstart](QUICKSTART_10_MIN.md)
- [Operations Notes](OPERATIONS.md)
- [Project Roadmap](ROADMAP.md)
