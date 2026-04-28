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

A [PEP 249 (DB-API 2.0)](https://peps.python.org/pep-0249/) compatible driver that lets you query `.xlsx` workbooks with SQL — no database server required.

> **This is not a document-preservation tool.** excel-dbapi treats worksheets as
> row-oriented datasets. It does not guarantee preservation of Excel formatting,
> charts, images, conditional formatting, or formulas — particularly with the
> pandas engine, which rewrites the entire workbook on save. If you need a
> round-trip-safe Excel editor, use openpyxl directly.

### ✅ Best for

- Querying small-to-medium `.xlsx` workbooks (up to ~50k rows) with SQL
- Automating Excel read/write with a familiar DB-API interface
- Teaching SQL without database setup
- Prototyping data pipelines before moving to a real database

### ❌ Not for

- Large datasets (50k+ rows) — use SQLite, DuckDB, or PostgreSQL
- Preserving Excel formatting, charts, or macros through write cycles
- Concurrent multi-writer scenarios
- Production OLTP/OLAP workloads

## Documentation

- **[SQL Specification](docs/SQL_SPEC.md)** — authoritative feature matrix and SQL subset reference (v1.0)
- [Usage Guide](docs/USAGE.md) — engine comparison, configuration, advanced patterns
- [Engine Selection Guide](docs/engines.md) — choosing the right backend
- [10-Minute Quickstart](docs/QUICKSTART_10_MIN.md)
- [Engine Benchmarks](docs/BENCHMARKS.md) — row limits, performance characteristics, preservation matrix
- [Microsoft Graph Backend](docs/graph-backend.md) — remote Excel on OneDrive/SharePoint
- [Project Roadmap](docs/ROADMAP.md)
- [Development Guide](docs/DEVELOPMENT.md)
- [Operations Notes](docs/OPERATIONS.md)

## Limitations

Before you begin, understand what excel-dbapi is **not**:

- **Not full SQL** — a documented SQL subset (see [SQL Specification](docs/SQL_SPEC.md))
- **Not a document-preservation tool** — the pandas engine drops all formatting, charts, images, and formulas on save; openpyxl preserves most formatting but some Excel features (e.g. conditional formatting rules, data validation, sparklines) may not survive round-trips through SQL DML
- **No concurrent writes** — single-writer model; advisory PID-based file locking prevents accidental double-opens within the same machine, but it is process-local and **not** a distributed lock
- **Not for large datasets** — designed for worksheets up to ~50k rows; beyond that, use a real database
- **No transactional rollback guarantees** — rollback restores an in-memory snapshot, not a WAL-based recovery; a crash mid-save can lose data
- **Identifier grammar** — both unquoted Unicode identifiers (`이름`, `naïve`) and double-quoted identifiers (`"Full Name"`, `"이름"`) are supported

If you need relational guarantees, concurrent access, or large-scale data, use SQLite or PostgreSQL.

## SQL Feature Set (Stable)

All features below are **Stable** — covered by the [SQL Spec v1.0](docs/SQL_SPEC.md) contract and will not have breaking changes.

- `SELECT` with aliases, arithmetic/CASE expressions, `DISTINCT`, `WHERE`, `GROUP BY`, `HAVING`, `ORDER BY`, `LIMIT`, `OFFSET`
- JOINs: `INNER`, `LEFT`, `RIGHT`, `FULL OUTER`, `CROSS`
- Aggregates: `COUNT`, `SUM`, `AVG`, `MIN`, `MAX`, `COUNT(DISTINCT col)`, `FILTER (WHERE ...)`
- Scalar functions: `UPPER`, `LOWER`, `LENGTH`, `TRIM`, `SUBSTR`, `COALESCE`, `NULLIF`, `CONCAT`, `ABS`, `ROUND`, `REPLACE`, `YEAR`, `MONTH`, `DAY`
- `CAST(expr AS type)` — `INTEGER`, `REAL`, `TEXT`, `DATE`, `DATETIME`, `BOOLEAN`
- Subqueries: `WHERE col [NOT] IN (SELECT ...)`, `EXISTS (SELECT ...)`, scalar subqueries
- Set operations: `UNION`, `UNION ALL`, `INTERSECT`, `EXCEPT`
- DML: `INSERT` (single/multi-row, `INSERT ... SELECT`), UPSERT (`ON CONFLICT`), `UPDATE`, `DELETE`
- DDL: `CREATE TABLE`, `DROP TABLE`, `ALTER TABLE ADD/DROP/RENAME COLUMN`
- Parameters: `?` positional placeholders (qmark paramstyle)

### Experimental Features

These features are implemented but may change semantics or be removed in a future release:

- **Window functions** (`ROW_NUMBER`, `RANK`, `DENSE_RANK`, running aggregates) — [SQL Spec § Window Functions](docs/SQL_SPEC.md)
- **CTEs** (`WITH ... AS`) — non-recursive only — [SQL Spec § CTEs](docs/SQL_SPEC.md)

For the full feature matrix with per-feature notes, see [docs/SQL_SPEC.md § Authoritative Feature Matrix](docs/SQL_SPEC.md#2-authoritative-feature-matrix).

---

## Who is this for?

- **Data analysts** who want to query Excel files with SQL instead of manual filtering
- **Citizen developers** automating small workflows with familiar SQL syntax
- **Educators** teaching SQL concepts without setting up a database
- **Prototypers** building quick data pipelines before moving to a real database

---

## Installation

```bash
pip install excel-dbapi
```

See [CHANGELOG](CHANGELOG.md) for release history.

---

## Quick Start

```python
from excel_dbapi import connect

# Open an Excel file and query it
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())
```

### Insert, Update, Delete

```python
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()

    # Insert with parameter binding (recommended)
    cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)", (1, "Alice"))

    # Update
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")

    # Delete
    cursor.execute("DELETE FROM Sheet1 WHERE id = 2")
```

### Multi-row Insert

```python
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()

    # Insert multiple rows at once
    cursor.execute("INSERT INTO Sheet1 VALUES (1, 'Alice'), (2, 'Bob'), (3, 'Carol')")

    # INSERT...SELECT: copy rows from another sheet
    cursor.execute("INSERT INTO Sheet2 (id, name) SELECT id, name FROM Sheet1 WHERE id > 1")
```

### Create and Drop Sheets

```python
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE NewSheet (id, name)")
    cursor.execute("DROP TABLE NewSheet")
```

---

## Engine Options

| Engine | Description | Dependency | Preserves formatting |
|--------|-------------|------------|---------------------|
| openpyxl (default) | Cell-level read/write | openpyxl | ✅ Yes |
| pandas | DataFrame-based operations | pandas, openpyxl | ❌ **No** — rewrites entire workbook |
| graph | Microsoft Graph API (remote) | httpx | ✅ (cell values only) |

```python
from excel_dbapi import connect

conn = connect("sample.xlsx", engine="openpyxl")  # default
conn = connect("sample.xlsx", engine="pandas")
```

### Engine Capability Matrix

| Capability | openpyxl | pandas | graph |
|---|---|---|---|
| Read support | ✅ | ✅ | ✅ |
| Write support | ✅ | ✅ | ✅ (opt-in, `readonly=False`) |
| Preserves formatting/charts/images | ✅ | ❌ (rewrites workbook) | ✅ (updates cell values only) |
| Transactions (commit/rollback) | ✅ (in-memory snapshot) | ✅ (in-memory snapshot) | ❌ (writes are immediate) |
| `data_only=False` (read formulas) | ✅ | ❌ | ❌ |
| File locking | ✅ (advisory PID-based) | ✅ (advisory PID-based) | N/A (remote) |
| Remote/cloud access | ❌ | ❌ | ✅ (Microsoft Graph) |
| `.workbook` access | ✅ | ❌ | ❌ |
| Formula injection defense | ✅ (default on) | ✅ (default on) | ✅ (default on) |

Choose **openpyxl** (default) for local files where you need formatting preservation and formula access.
Choose **pandas** when you prefer DataFrame-based workflows and don't need formatting.
Choose **graph** for remote Excel on OneDrive/SharePoint via Microsoft Graph API.

For detailed engine comparison and benchmarks, see the [Engine Selection Guide](docs/engines.md) and [Benchmarks](docs/BENCHMARKS.md).

---

## WHERE Operators

| Operator | Example | Description |
|----------|---------|-------------|
| `=`, `!=`, `<>` | `WHERE id = 1` | Equality / inequality |
| `>`, `>=`, `<`, `<=` | `WHERE score >= 80` | Comparison |
| `IS NULL` / `IS NOT NULL` | `WHERE name IS NOT NULL` | NULL checks |
| `IN` | `WHERE name IN ('Alice', 'Bob')` | Set membership |
| `BETWEEN` | `WHERE score BETWEEN 70 AND 90` | Inclusive range |
| `LIKE` / `ILIKE` | `WHERE name LIKE 'A%'` | Pattern matching (ILIKE = case-insensitive) |
| `NOT LIKE` / `NOT ILIKE` | `WHERE name NOT LIKE 'A%'` | Negated pattern matching |
| `NOT IN` | `WHERE id NOT IN (1, 2)` | Negated set membership |
| `NOT BETWEEN` | `WHERE x NOT BETWEEN 1 AND 5` | Negated range |
| `AND` / `OR` / `NOT` | `WHERE x = 1 AND y = 2` | Logical connectives |

> **NULL semantics**: Comparisons with NULL follow SQL three-valued logic (TRUE / FALSE / UNKNOWN). `WHERE x = NULL` returns no rows; use `IS NULL` instead.

**LIKE patterns:** `%` matches any sequence of characters, `_` matches any single character.

```python
from excel_dbapi import connect

with connect("sample.xlsx") as conn:
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

### Compound Queries (Set Operations)

```python
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()

    cursor.execute("SELECT id FROM t1 UNION SELECT id FROM t2")
    cursor.execute("SELECT id FROM t1 UNION ALL SELECT id FROM t2")
    cursor.execute("SELECT id FROM t1 INTERSECT SELECT id FROM t2")
    cursor.execute("SELECT id FROM t1 EXCEPT SELECT id FROM t2")
```

---

## Safety Defaults

### Formula Injection Defense

By default, `excel-dbapi` sanitizes cell values on write (INSERT/UPDATE) to prevent
[formula injection attacks](https://owasp.org/www-community/attacks/CSV_Injection).
Strings starting with `=`, `+`, `-`, `@`, `\t`, or `\r` are automatically prefixed
with a single quote (`'`) so they are stored as plain text, not executed as formulas.

```python
from excel_dbapi import connect

# Default: sanitization ON (recommended)
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
                   (1, "=SUM(A1:A10)"))
    # Stored as: '=SUM(A1:A10)  (safe, not executed as formula)

# Opt out if you intentionally write formulas
with connect("sample.xlsx", sanitize_formulas=False) as conn:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Sheet1 (id, formula) VALUES (?, ?)",
                   (1, "=SUM(A1:A10)"))
    # Stored as: =SUM(A1:A10)  (executed as formula in Excel)
```

---

## Transactions

```python
from excel_dbapi import connect

with connect("sample.xlsx", autocommit=False) as conn:
    cursor = conn.cursor()
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    conn.rollback()
```

> **Important:** Transactions use in-memory snapshots, not write-ahead logging.
> `rollback()` restores the last committed snapshot in memory — it does not undo
> writes already flushed to disk. When `autocommit=True` (the default), each
> write is saved immediately and `rollback()` is not supported. The graph engine
> does not support transactions at all; writes are applied immediately to the
> remote workbook.
>
> **Context manager behaviour:** `with connect(...) as conn:` calls `conn.close()`
> on exit — it does **not** auto-commit or auto-rollback. If you need to persist
> changes when `autocommit=False`, call `conn.commit()` explicitly before the
> `with` block ends.

## Cursor Metadata

```python
from excel_dbapi import connect

with connect("sample.xlsx") as conn:
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

**Fix:** Check the sheet name spelling. Sheet names are resolved case-insensitively.

### PandasEngine drops formatting

`PandasEngine` reads data into a DataFrame and writes it back. This process drops
Excel formatting, charts, images, and formulas.

**Fix:** Use the default `openpyxl` engine if you need to preserve formatting.

### Integer vs. string comparison (Pandas)

The Pandas engine preserves Python types. If a column contains integers,
`WHERE id = '2'` (string) won't match — use `WHERE id = 2` (no quotes).

**Fix:** Omit quotes around numeric values in WHERE clauses when using the Pandas engine.

---

## Experimental: Remote Excel via Microsoft Graph API

> **Status**: Experimental — API may change in future releases.

excel-dbapi can access remote Excel files on OneDrive/SharePoint via the Microsoft Graph API.

Supported Graph DSNs are ID-based:

- `msgraph://drives/{drive_id}/items/{item_id}`
- `sharepoint://sites/{site_name}/drives/{drive_id}/items/{item_id}`
- `onedrive://me/drive/items/{item_id}`

```bash
pip install excel-dbapi[graph]
```

```python
from excel_dbapi import connect

conn = connect(
    "msgraph://drives/{drive_id}/items/{item_id}",
    engine="graph",
    credential=your_credential,
    autocommit=True,
)
cursor = conn.cursor()
cursor.execute("SELECT * FROM Sheet1")
print(cursor.fetchall())
conn.close()
```

The Graph backend is **read-only by default**. Write operations require explicit opt-in
and a credential/token provider with appropriate Graph API permissions.

Graph metadata sync is best-effort for write operations: if worksheet mutation succeeds
but metadata sync fails, excel-dbapi keeps the worksheet change and logs a warning.

For DSN formats and dependency choices, see the
[Graph Backend Guide](docs/graph-backend.md).

---

## Related Projects

- [sqlalchemy-excel](https://github.com/yeongseon/sqlalchemy-excel) — SQLAlchemy dialect that uses excel-dbapi as its DB-API 2.0 driver. Use `create_engine("excel:///file.xlsx")` for full ORM support.

---

## Examples

- `examples/basic_usage.py`
- `examples/write_operations.py`
- `examples/transactions.py`
- `examples/advanced_query.py`
- `examples/pandas_engine.py`

---

## License

MIT License
