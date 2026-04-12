<p align="left">
  <img src="https://raw.githubusercontent.com/yeongseon/excel-dbapi/main/logo.svg" alt="excel-dbapi" width="48" height="48" align="middle" />
  <strong style="font-size: 2em;">excel-dbapi</strong>
</p>

![CI](https://github.com/yeongseon/excel-dbapi/actions/workflows/ci.yml/badge.svg)
[![codecov](https://codecov.io/gh/yeongseon/excel-dbapi/branch/main/graph/badge.svg)](https://codecov.io/gh/yeongseon/excel-dbapi)
[![PyPI](https://img.shields.io/pypi/v/excel-dbapi.svg)](https://pypi.org/project/excel-dbapi/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Docs](https://img.shields.io/badge/docs-GitHub-blue.svg)](https://github.com/yeongseon/excel-dbapi/tree/main/docs)

A lightweight, Python DB-API 2.0 compliant connector for Excel files.
Use SQL to query, insert, update, and delete rows in `.xlsx` workbooks — no database server required.

## Who is this for?

- **Data analysts** who want to query Excel files with SQL instead of manual filtering
- **Citizen developers** automating small workflows with familiar SQL syntax
- **Educators** teaching SQL concepts without setting up a database
- **Prototypers** building quick data pipelines before moving to a real database

### Who is this NOT for?

- If you need JOINs, GROUP BY, subqueries, or advanced SQL → use SQLite or PostgreSQL
- If you need concurrent writes from multiple processes → use a real database
- If your Excel file has 100k+ rows → use pandas directly or a database

---

## Features

- Python DB-API 2.0 compliant interface (PEP 249)
- Query Excel files using SQL syntax
- Supports SELECT, INSERT, UPDATE, DELETE
- Basic DDL support (CREATE TABLE, DROP TABLE)
- WHERE conditions with AND/OR and comparison operators
- IN, BETWEEN, LIKE operators in WHERE clauses
- ORDER BY and LIMIT for SELECT
- Sheet-to-Table mapping
- Pandas & Openpyxl engine selector
- Formula injection defense (enabled by default)
- Transaction simulation (commit/rollback)

---

## Installation

```bash
pip install excel-dbapi
```

See [CHANGELOG](CHANGELOG.md) for release history.

---

## Quick Start

```python
from excel_dbapi.connection import ExcelConnection

# Open an Excel file and query it
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())
```

### Insert, Update, Delete

```python
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()

    # Insert with parameter binding (recommended)
    cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)", (1, "Alice"))

    # Update
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")

    # Delete
    cursor.execute("DELETE FROM Sheet1 WHERE id = 2")
```

### Create and Drop Sheets

```python
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE NewSheet (id, name)")
    cursor.execute("DROP TABLE NewSheet")
```

### Engine Options

| Engine | Description | Dependency |
|--------|-------------|------------|
| openpyxl (default) | Fast sheet access | openpyxl |
| pandas | DataFrame-based operations | pandas, openpyxl |

```python
conn = ExcelConnection("sample.xlsx", engine="openpyxl")  # default
conn = ExcelConnection("sample.xlsx", engine="pandas")
```

### WHERE Operators

| Operator | Example | Description |
|----------|---------|-------------|
| `=`, `!=`, `<>` | `WHERE id = 1` | Equality / inequality |
| `>`, `>=`, `<`, `<=` | `WHERE score >= 80` | Comparison |
| `IS NULL` / `IS NOT NULL` | `WHERE name IS NOT NULL` | NULL checks |
| `IN` | `WHERE name IN ('Alice', 'Bob')` | Set membership |
| `BETWEEN` | `WHERE score BETWEEN 70 AND 90` | Inclusive range |
| `LIKE` | `WHERE name LIKE 'A%'` | Pattern matching |
| `AND` / `OR` | `WHERE x = 1 AND y = 2` | Logical connectives |

**LIKE patterns:** `%` matches any sequence of characters, `_` matches any single character.

```python
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()

    # IN operator
    cursor.execute("SELECT * FROM Sheet1 WHERE name IN ('Alice', 'Bob')")

    # BETWEEN operator
    cursor.execute("SELECT * FROM Sheet1 WHERE score BETWEEN 70 AND 90")

    # LIKE operator
    cursor.execute("SELECT * FROM Sheet1 WHERE name LIKE 'A%'")

    # All operators support parameter binding
    cursor.execute("SELECT * FROM Sheet1 WHERE name IN (?, ?)", ("Alice", "Bob"))
    cursor.execute("SELECT * FROM Sheet1 WHERE score BETWEEN ? AND ?", (70, 90))
    cursor.execute("SELECT * FROM Sheet1 WHERE name LIKE ?", ("A%",))
```

---

## Safety Defaults

### Formula Injection Defense

By default, `excel-dbapi` sanitizes cell values on write (INSERT/UPDATE) to prevent
[formula injection attacks](https://owasp.org/www-community/attacks/CSV_Injection).
Strings starting with `=`, `+`, `-`, `@`, `\t`, or `\r` are automatically prefixed
with a single quote (`'`) so they are stored as plain text, not executed as formulas.

```python
# Default: sanitization ON (recommended)
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
                   (1, "=SUM(A1:A10)"))
    # Stored as: '=SUM(A1:A10)  (safe, not executed as formula)

# Opt out if you intentionally write formulas
with ExcelConnection("sample.xlsx", sanitize_formulas=False) as conn:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Sheet1 (id, formula) VALUES (?, ?)",
                   (1, "=SUM(A1:A10)"))
    # Stored as: =SUM(A1:A10)  (executed as formula in Excel)
```

---

## Transaction Example

```python
with ExcelConnection("sample.xlsx", autocommit=False) as conn:
    cursor = conn.cursor()
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    conn.rollback()
```

When autocommit is enabled, `rollback()` is not supported.

## Cursor Metadata

```python
with ExcelConnection("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM Sheet1")
    print(cursor.description)
    print(cursor.rowcount)
```

---

## Troubleshooting

### "Column 'xyz' not found"

The column name in your SQL doesn't match any header in the sheet.

```
ProgrammingError: Column 'nmae' not found in Sheet1. Available columns: ['id', 'name', 'email']
```

**Fix:** Check the spelling. Column names must match the first row (header) of the sheet exactly.

### "Table 'SheetX' not found"

The sheet name in your SQL doesn't match any sheet in the workbook.

```
ProgrammingError: Table 'Shee1' not found. Available sheets: ['Sheet1', 'Sheet2']
```

**Fix:** Check the sheet name spelling. Use the exact sheet name (case-sensitive) shown in your Excel file.

### PandasEngine drops formatting

`PandasEngine` reads data into a DataFrame and writes it back. This process drops
Excel formatting, charts, images, and formulas.

**Fix:** Use the default `openpyxl` engine if you need to preserve formatting.

### Integer vs. string comparison (Pandas)

The Pandas engine preserves Python types. If a column contains integers,
`WHERE id = '2'` (string) won't match — use `WHERE id = 2` (no quotes).

**Fix:** Omit quotes around numeric values in WHERE clauses when using the Pandas engine.

---

## Limitations and Operational Guidance

- `PandasEngine` rewrites workbooks and may drop formatting, charts, and formulas.
- `OpenpyxlEngine` loads with `data_only=True`, so formulas are evaluated to values when reading.
- Use a **single-writer model** for writes. Avoid writing to the same file from multiple processes.
- Save is implemented with a temporary file + atomic replace (`os.replace`) for safer persistence.
- No support for JOIN, GROUP BY, HAVING, or subqueries.

## Roadmap

- Remote file connection improvements

See [Project Roadmap](docs/ROADMAP.md) for details.

---


## Related Projects

- [sqlalchemy-excel](https://github.com/yeongseon/sqlalchemy-excel) — SQLAlchemy dialect that uses excel-dbapi as its DB-API 2.0 driver. Use `create_engine("excel:///file.xlsx")` for full ORM support.
---

## Documentation

- [Usage Guide](docs/USAGE.md)
- [Development Guide](docs/DEVELOPMENT.md)
- [Project Roadmap](docs/ROADMAP.md)
- [10-Minute Quickstart](docs/QUICKSTART_10_MIN.md)
- [Operations Notes](docs/OPERATIONS.md)
- [Public Roadmap](docs/PUBLIC_ROADMAP.md)

## Examples

- `examples/basic_usage.py`
- `examples/write_operations.py`
- `examples/transactions.py`
- `examples/advanced_query.py`
- `examples/pandas_engine.py`

---

## License

MIT License
