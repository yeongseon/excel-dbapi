# Engine Benchmarks & Characteristics

> Version: excel-dbapi 0.4.x series
> Last updated: 2026-04-15

This document provides quantitative and qualitative characteristics for each
backend engine. Use it to choose the right engine for your workload and to
understand operational limits.

---

## 1. Engine Overview

| Property | openpyxl (default) | pandas | graph |
|---|---|---|---|
| Storage model | In-memory openpyxl `Workbook` object | In-memory `pandas.DataFrame` per sheet | Remote Microsoft Graph REST API |
| Read source | `openpyxl.load_workbook()` | `pd.read_excel()` | Graph `usedRange` endpoint |
| Write mechanism | Cell-level in-place mutation → atomic `os.replace()` | `pd.ExcelWriter` rebuilds entire workbook → atomic `os.replace()` | PATCH individual ranges via Graph API |
| Dependencies | `openpyxl` | `pandas`, `openpyxl` | `httpx` |

---

## 2. Row & Memory Limits

excel-dbapi is designed for worksheet-scale data — typically up to **~50,000 rows**.
Beyond that threshold, consider migrating to SQLite or PostgreSQL.

### Configurable Guards

Both `max_rows` and `max_memory_mb` are configurable per-connection via
`backend_options`:

```python
conn = ExcelConnection(
    "data.xlsx",
    backend_options={"max_rows": 10_000, "max_memory_mb": 100.0},
)
```

| Guard | Default | Behavior |
|---|---|---|
| `max_rows` | None (no limit) | Warning at 80% of limit; `OperationalError` when exceeded |
| `max_memory_mb` | None (no limit) | Warning at 80% of limit; `OperationalError` when exceeded |

### Practical Row-Count Guidance

| Row Range | Expectation |
|---|---|
| < 10,000 | Fast on all engines. Sub-second reads/writes. |
| 10,000–50,000 | Workable. openpyxl and pandas handle this well for typical column counts (< 50 columns). Graph latency depends on network. |
| 50,000–100,000 | Possible but slow. Memory usage grows linearly. Snapshot/rollback operations duplicate the entire dataset in memory. |
| > 100,000 | Not recommended. Use a real database. |

### Memory Characteristics

| Engine | Memory Model | Snapshot Cost |
|---|---|---|
| openpyxl | Entire workbook held in memory as openpyxl objects (cells, styles, charts). Memory per row depends on column count and cell formatting. | `snapshot()` serializes the full workbook to a `BytesIO` buffer, then `restore()` reloads it. Cost ≈ 2× workbook memory during snapshot. |
| pandas | Each sheet is a `DataFrame`. Memory per row depends on column types and pandas dtype inference. Pending rows are buffered separately until flushed. | `snapshot()` deep-copies all DataFrames. Cost ≈ 2× data memory. |
| graph | No local data cache — each `read_sheet()` fetches from Graph API. Row/memory limits are checked on the fetched response. | `snapshot()` returns `None` (no-op). `restore()` closes the session and clears caches. No memory duplication. |

---

## 3. Write Performance Characteristics

### openpyxl

- **Write strategy**: Modifies cells in-place on the openpyxl `Worksheet` object,
  preserving formatting. Surplus rows/columns are deleted.
- **Save**: Writes to a temporary file, then atomically replaces the target with
  `os.replace()`. Safe against partial writes on crash.
- **Append**: Uses `ws.append()` — fast for bulk inserts.
- **Cost**: Proportional to the number of modified cells, not total sheet size.

### pandas

- **Write strategy**: Replaces the entire `DataFrame` for the sheet. On save,
  **all sheets** are rewritten from scratch via `pd.ExcelWriter`.
- **Save**: Writes to a temporary file, then atomically replaces the target.
- **Append**: Buffers rows in a pending list; flushes with `pd.concat()` on
  next read or save.
- **Cost**: Full workbook rewrite on every `save()`. Scales with total data
  volume across all sheets, not just the modified sheet.

### graph

- **Write strategy**: Targeted PATCH requests for changed row ranges. If less
  than 50% of rows changed, only changed rows are patched. Otherwise, a full
  range rewrite is performed.
- **DELETE optimization**: When rows are removed, uses Graph range `delete`
  with shift-up instead of rewriting.
- **Append**: Single PATCH to the next empty row.
- **Cost**: Dominated by HTTP round-trips. Each row group is a separate request.
  Latency depends on network conditions and Graph API throttling.

---

## 4. Formatting & Data Preservation Matrix

This matrix shows what survives a round-trip (read → SQL DML → write → reopen)
through each engine.

| Feature | openpyxl | pandas | graph |
|---|---|---|---|
| **Cell values** (strings, numbers, dates) | ✅ Preserved | ✅ Preserved | ✅ Preserved |
| **Cell formatting** (fonts, borders, fills, alignment) | ✅ Preserved (in-place write) | ❌ **Dropped** — workbook rebuilt from DataFrames | ✅ Preserved — only cell values are updated |
| **Number formats** (currency, percentage, date) | ✅ Preserved | ❌ **Dropped** | ✅ Preserved |
| **Conditional formatting** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Data validation** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Charts** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Images / drawings** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Comments / notes** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Formulas** | ⚠️ With `data_only=False`: formula text readable; with `data_only=True` (default): cached values only | ❌ **Not supported** — `data_only=False` raises `NotSupportedError` | ❌ **Not supported** — `data_only=False` raises `NotSupportedError` |
| **Merged cells** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Freeze panes** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Print settings** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Named ranges** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Sheet tab colors** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |
| **Column widths / row heights** | ✅ Preserved | ❌ **Dropped** | ✅ Preserved |
| **Hyperlinks** | ✅ Preserved by openpyxl | ❌ **Dropped** | ✅ Preserved |

**Key takeaway**: The pandas engine destroys all non-data content on save. If you
need to preserve anything beyond raw cell values, use openpyxl (local) or graph
(remote).

> **Note on openpyxl preservation**: openpyxl's preservation is subject to
> openpyxl's own format support. Some advanced Excel features (VBA macros,
> Power Query, pivot tables, sparklines) may not survive a load/save cycle
> even through openpyxl alone. This is an openpyxl limitation, not an
> excel-dbapi limitation.

---

## 5. Transaction & Concurrency Model

| Capability | openpyxl | pandas | graph |
|---|---|---|---|
| `autocommit=True` (default) | Each write saves to disk immediately | Each write saves to disk immediately | Writes are always immediate |
| `autocommit=False` | Writes accumulate in memory; `commit()` saves to disk; `rollback()` restores snapshot | Same as openpyxl | ❌ `NotSupportedError` |
| Snapshot mechanism | Serialize workbook to `BytesIO` | Deep-copy all DataFrames | No-op (no local state) |
| Rollback guarantee | In-memory only — crash during `save()` loses data | In-memory only | N/A |
| Concurrent readers | ✅ Multiple processes can read simultaneously | ✅ Multiple processes can read simultaneously | ✅ Multiple sessions can read |
| Concurrent writers | ❌ Single-writer model; advisory PID-based `.lock` file | ❌ Single-writer model; advisory PID-based `.lock` file | ⚠️ ETag-based optimistic concurrency (`fail` or `force` strategy) |
| File locking | Advisory PID-based `.lock` file; stale lock auto-cleared | Advisory PID-based `.lock` file; stale lock auto-cleared | N/A (remote; uses ETag) |

---

## 6. Feature Support Comparison

| Feature | openpyxl | pandas | graph |
|---|---|---|---|
| Read `.xlsx` | ✅ | ✅ | ✅ (remote) |
| Write `.xlsx` | ✅ | ✅ | ✅ (opt-in `readonly=False`) |
| `data_only=False` (formulas) | ✅ | ❌ | ❌ |
| `get_workbook()` | ✅ (returns openpyxl `Workbook`) | ❌ | ❌ |
| `sanitize_formulas` | ✅ | ✅ | ✅ |
| `max_rows` guard | ✅ | ✅ | ✅ |
| `max_memory_mb` guard | ✅ | ✅ | ✅ |
| Atomic file save | ✅ (temp file + `os.replace()`) | ✅ (temp file + `os.replace()`) | N/A (remote API) |
| Session management | N/A | N/A | ✅ (auto-reopen on expiry) |

---

## 7. When to Use Which Engine

| Scenario | Recommended Engine | Reason |
|---|---|---|
| Local file, preserve formatting | **openpyxl** | Only local engine that preserves all non-data content |
| Local file, formula read/write | **openpyxl** (`data_only=False`) | Only engine that supports formula access |
| Data pipeline with DataFrames | **pandas** | Direct DataFrame integration; formatting not needed |
| Remote Excel on OneDrive/SharePoint | **graph** | Only engine that supports remote access |
| Teaching or prototyping | **openpyxl** | Simplest setup, no extra dependencies |
| Bulk data import (local) | **openpyxl** or **pandas** | Both handle bulk inserts; pandas may be faster for very large DataFrames due to vectorized operations |
| Read-only analytics on remote files | **graph** (`readonly=True`) | No write session overhead |

---

## 8. Known Limitations by Engine

### openpyxl
- VBA macros are not preserved through load/save cycles (openpyxl limitation).
- Pivot tables and Power Query connections are dropped (openpyxl limitation).
- Sparklines are not preserved (openpyxl limitation).
- Large workbooks with heavy formatting consume significant memory.

### pandas
- **All formatting, charts, images, comments, and formulas are destroyed** on
  every save — this is inherent to the DataFrame-based storage model.
- Type coercion follows pandas rules: integers in columns with `None` values
  become floats (`NaN`); date parsing may alter original formats.
- `WHERE id = '2'` (string) will not match an integer column — use `WHERE id = 2`.

### graph
- Path-based DSNs (`sharepoint://...path/to/file.xlsx`) are not implemented;
  only ID-based DSNs are supported.
- Write operations require explicit `readonly=False` and appropriate Graph API
  permissions.
- No transaction support — writes are immediately persisted and cannot be rolled back.
- Metadata sync after DDL is best-effort: if mutation succeeds but metadata sync
  fails, the data change remains and a warning is logged.
- Subject to Microsoft Graph API rate limits and throttling.
