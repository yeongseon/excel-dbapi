# Usage Guide

## Installation

```bash
pip install excel-dbapi
```

## Graph Backend DSN and Installation

Use the Graph backend to query remote workbooks on Microsoft 365.

### Extras

- `pip install excel-dbapi[graph]`: Graph backend with generic token provider support.
- `pip install excel-dbapi[graph-azure]`: includes Azure Identity for `DefaultAzureCredential` and other Azure credential flows.

Choose `graph-azure` when you want excel-dbapi to acquire tokens through `azure-identity` directly.

### DSN Formats

- `msgraph://drives/{drive_id}/items/{item_id}`
  - Generic Graph endpoint form when you already know drive/item IDs.
- `sharepoint://sites/{site_name}/drives/{drive_id}/items/{item_id}`
  - SharePoint workbook by site/drive/item IDs.
- `onedrive://me/drive/items/{item_id}`
  - Signed-in user's OneDrive workbook by item ID.

Path-based forms such as `sharepoint://.../Shared Documents/path/to/workbook.xlsx`
and `onedrive://path/to/workbook.xlsx` are not implemented.

### Example

```python
from excel_dbapi import connect

with connect(
    "msgraph://drives/{drive_id}/items/{item_id}",
    engine="graph",
    credential=your_credential,
) as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())
```

### Graph Write-Sync Limitation

Graph is a non-transactional backend. If a worksheet mutation succeeds but metadata
sync fails, excel-dbapi keeps the workbook change and logs a warning instead of
rolling back the mutation.

## Basic Example

```python
from excel_dbapi import connect

# Open a connection to the Excel file
with connect("sample.xlsx") as conn:
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

### Notes

JOINs, GROUP BY/HAVING, DISTINCT, OFFSET, subqueries, CTEs, and aggregate functions
are supported. See [SQL_SPEC.md](SQL_SPEC.md) for precise behavior and edge-case
limits.

## Write Operations

```python
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (1, 'Alice')")
    cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
    cursor.execute("DELETE FROM Sheet1 WHERE id = 2")
```

## Bulk Inserts

```python
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.executemany(
        "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
        [(2, "Bob"), (3, "Cara")],
    )
```

## DDL

```python
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE NewSheet (id, name)")
    cursor.execute("DROP TABLE NewSheet")
```

## Transactions

Autocommit is enabled by default. To use manual transactions:

```python
with connect("sample.xlsx", autocommit=False) as conn:
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
with connect("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, name FROM Sheet1 WHERE id >= ? ORDER BY id DESC LIMIT 1",
        (1,),
    )
    print(cursor.fetchall())
```

## Engine Comparison

excel-dbapi ships three backends. They share the same SQL interface but differ in
storage model, feature coverage, and operational trade-offs.

### Quick Reference

| Capability | openpyxl (default) | pandas | graph |
|---|---|---|---|
| **Read** | ✅ local `.xlsx` | ✅ local `.xlsx` | ✅ remote (Microsoft Graph) |
| **Write** | ✅ | ✅ | ✅ opt-in (`readonly=False`) |
| **Preserves formatting** | ✅ | ❌ rewrites workbook | ✅ updates values only |
| **Transactions** | ✅ commit / rollback (in-memory snapshot) | ✅ commit / rollback (in-memory snapshot) | ❌ writes are immediate |
| **`data_only=False`** | ✅ read raw formulas | ❌ raises `NotSupportedError` | ❌ raises `NotSupportedError` |
| **File locking** | ✅ advisory PID-based `.lock` file | ✅ advisory PID-based `.lock` file | N/A (remote; uses ETag concurrency) |
| **`.workbook`** | ✅ returns openpyxl `Workbook` | ❌ raises `NotSupportedError` | ❌ |
| **Remote access** | ❌ local only | ❌ local only | ✅ OneDrive / SharePoint |
| **Formula injection defense** | ✅ on by default | ✅ on by default | ✅ on by default |
| **Dependency** | `openpyxl` | `pandas`, `openpyxl` | `httpx` |

### openpyxl (default)

The openpyxl backend loads a `.xlsx` workbook via `openpyxl.load_workbook()` and
operates on the live `Workbook` object. It preserves formatting, charts, images, and
comments through load/save cycles (subject to openpyxl's own format support).

- **Atomic saves**: writes to a temp file then replaces the target with `os.replace()`.
- **Snapshot rollback**: `snapshot()` serialises the workbook into a `BytesIO` buffer;
  `restore()` reloads from the buffer. This is an in-memory snapshot, not a WAL.
- **Formula access**: set `data_only=False` on the connection to read formula text
  instead of cached values.
- **Direct workbook access**: `connection.workbook` returns the openpyxl `Workbook`
  for direct styling, data-validation, or chart manipulation.
- **Best for**: local workflows that need formatting preservation, formula access, or
  direct workbook manipulation.

### pandas

The pandas backend reads all sheets into `pandas.DataFrame` objects via `pd.read_excel()`
and writes them back with `pd.ExcelWriter` (engine=`"openpyxl"`).

- **Workbook rewrite**: every `save()` rebuilds the workbook from DataFrames.
  **Formatting, charts, images, comments, and formulas are dropped.**
- **No formula access**: `data_only=False` raises `NotSupportedError`.
- **No `.workbook` access**: `.workbook` raises `NotSupportedError` because
  there is no persistent openpyxl `Workbook` object.
- **Type fidelity**: pandas preserves Python types on read. `WHERE id = '2'`
  (string) will not match an integer column — use `WHERE id = 2`.
- **Best for**: DataFrame-centric pipelines where you do not need formatting or formulas.

### graph

The graph backend accesses remote Excel workbooks on OneDrive / SharePoint via the
Microsoft Graph API.

- **Read-only by default**: pass `readonly=False` via backend options to enable writes.
- **Immediate persistence**: writable sessions use `persistChanges=true`. Changes are
  applied to the remote workbook immediately and **cannot be rolled back**.
- **Non-transactional**: `supports_transactions` is `False`. `autocommit=False` raises
  `NotSupportedError`. `rollback()` is not available.
- **No formula access**: `data_only=False` raises `NotSupportedError`.
- **Session management**: the backend opens a Graph workbook session and handles
  session expiry automatically (reopen + retry).
- **Concurrency**: uses ETag / `If-Match` optimistic concurrency and conflict
  strategies (`"fail"` or `"force"`).
- **Authentication**: requires a token provider — a static token string, a callable,
  an `azure-identity` credential, or a custom `TokenProvider` object. See
  [Graph Backend DSN and Installation](#graph-backend-dsn-and-installation).
- **Metadata sync**: best-effort. If metadata sync fails after a successful worksheet
  mutation, the workbook change is kept and a warning is logged.
- **Best for**: querying or updating remote Excel files on Microsoft 365.

### When to Use Which Engine

| Scenario | Recommended Engine |
|---|---|
| Local file, preserve formatting | openpyxl |
| Local file, formula read/write | openpyxl (`data_only=False`) |
| Data pipeline with DataFrames | pandas |
| Remote Excel on OneDrive/SharePoint | graph |
| Teaching or prototyping | openpyxl (simplest setup) |


## Limitations

- `PandasBackend` (engine=`"pandas"`) rewrites workbooks and may drop formatting, charts, and formulas.
- `OpenpyxlBackend` (engine=`"openpyxl"`) defaults to `data_only=True`, so formulas are read as cached values unless you set `data_only=False` via `connect(..., data_only=False)`.
- Some advanced SQL patterns remain limited; see [SQL_SPEC.md](SQL_SPEC.md) for the
  exact supported subset and restrictions.
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
with connect("sample.xlsx") as conn:
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
