# Operations Notes

## Concurrency model

`excel-dbapi` assumes a **single writer** per workbook path.

- Safe pattern: one process writes, many can read snapshots.
- Unsafe pattern: multiple concurrent writers to the same `.xlsx` file.

If multiple writers are required, coordinate with an external lock.

## Engine tradeoffs

### openpyxl engine (default)
- Better workbook fidelity for most classroom/business spreadsheets.
- Good default for mixed read/write workloads.

### pandas engine
- Convenient for DataFrame-heavy flows.
- Rewrites the workbook and may drop rich formatting, formulas, and charts.

## File integrity semantics

Saves are performed with a temp file followed by `os.replace(...)`.
This provides atomic replacement on supported filesystems and avoids partial writes in normal failure scenarios.
