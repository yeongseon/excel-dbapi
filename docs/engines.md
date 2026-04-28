# Engine Selection Guide

excel-dbapi ships three backends that share the same SQL interface but differ in
storage model, feature coverage, and operational trade-offs.

## Quick Decision

- **openpyxl** (default) — local `.xlsx` files; preserves formatting, supports formulas
- **pandas** — only if your pipeline is already DataFrame-centric; drops formatting on save
- **graph** — remote Excel on OneDrive/SharePoint via Microsoft Graph API

## Feature Matrix

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
| Dependency | `openpyxl` | `pandas`, `openpyxl` | `httpx` |

## openpyxl (default)

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

**Best for**: local workflows that need formatting preservation, formula access, or
direct workbook manipulation.

```python
from excel_dbapi import connect

with connect("sample.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())
```

## pandas

The pandas backend reads all sheets into `pandas.DataFrame` objects via `pd.read_excel()`
and writes them back with `pd.ExcelWriter` (engine=`"openpyxl"`).

- **Workbook rewrite**: every `save()` rebuilds the workbook from DataFrames.
  **Formatting, charts, images, comments, and formulas are dropped.**
- **No formula access**: `data_only=False` raises `NotSupportedError`.
- **No `.workbook` access**: raises `NotSupportedError` because there is no persistent
  openpyxl `Workbook` object.
- **Type fidelity**: pandas preserves Python types on read. `WHERE id = '2'`
  (string) will not match an integer column — use `WHERE id = 2`.

**Best for**: DataFrame-centric pipelines where you do not need formatting or formulas.

```python
from excel_dbapi import connect

with connect("sample.xlsx", engine="pandas") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")
    print(cursor.fetchall())
```

> **Note**: The pandas engine is an optional extra. Install
> with `pip install excel-dbapi[pandas]`.

## graph

The graph backend accesses remote Excel workbooks on OneDrive / SharePoint via the
Microsoft Graph API.

- **Read-only by default**: pass `readonly=False` via backend options to enable writes.
- **Immediate persistence**: writable sessions use `persistChanges=true`. Changes are
  applied to the remote workbook immediately and **cannot be rolled back**.
- **Non-transactional**: `autocommit=False` raises `NotSupportedError`.
- **Session management**: the backend opens a Graph workbook session and handles
  session expiry automatically (reopen + retry).
- **Concurrency**: uses ETag / `If-Match` optimistic concurrency and conflict
  strategies (`"fail"` or `"force"`).
- **Authentication**: requires a token provider — a static token string, a callable,
  an `azure-identity` credential, or a custom `TokenProvider` object.

**Best for**: querying or updating remote Excel files on Microsoft 365.

```python
from excel_dbapi import connect

conn = connect(
    "msgraph://drives/{drive_id}/items/{item_id}",
    engine="graph",
    credential=your_credential,
)
cursor = conn.cursor()
cursor.execute("SELECT * FROM Sheet1")
print(cursor.fetchall())
conn.close()
```

For full production deployment guidance, see the [Graph Backend Guide](graph-backend.md).

## When to Use Which

| Scenario | Recommended Engine |
|---|---|
| Local file, preserve formatting | openpyxl |
| Local file, formula read/write | openpyxl (`data_only=False`) |
| Data pipeline with DataFrames | pandas |
| Remote Excel on OneDrive/SharePoint | graph |
| Teaching or prototyping | openpyxl (simplest setup) |

## Further Reading

- [Usage Guide](USAGE.md) — configuration, advanced patterns
- [Engine Benchmarks](BENCHMARKS.md) — row limits, performance characteristics
- [SQL Specification](SQL_SPEC.md) — authoritative SQL subset reference
- [Graph Backend Guide](graph-backend.md) — Microsoft Graph setup and operations
