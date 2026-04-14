# Operations Notes

## Concurrency Model

`excel-dbapi` assumes a **single writer** per workbook path.

- Safe pattern: one process writes, many can read snapshots.
- Unsafe pattern: multiple concurrent writers to the same `.xlsx` file.

If multiple writers are required, coordinate with an external lock.

## Engine Tradeoffs

### openpyxl engine (default)

- Better workbook fidelity for most classroom/business spreadsheets.
- Good default for mixed read/write workloads.
- Defaults to `data_only=True` — formulas are read as cached values unless you configure `data_only=False` when creating the connection.

### pandas engine

- Convenient for DataFrame-heavy flows.
- Rewrites the workbook and may drop rich formatting, formulas, and charts.

## File Integrity Semantics

Saves are performed with a temp file followed by `os.replace(...)`.
This provides atomic replacement on supported filesystems and avoids partial writes in normal failure scenarios.

## Transaction Semantics

- **`autocommit=True`** (default): Changes are saved to disk after each `execute()`. `rollback()` raises `NotSupportedError`.
- **`autocommit=False`**: Changes are held in memory. `commit()` flushes to disk. `rollback()` restores the in-memory snapshot taken at connection open or last `commit()`.

This is **not** ACID — there is no write-ahead log or crash recovery.
