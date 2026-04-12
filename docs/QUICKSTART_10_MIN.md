# 10-Minute Quickstart

## 1) Install

```bash
pip install excel-dbapi
```

## 2) Copy/paste first query

```python
from excel_dbapi.connection import ExcelConnection

with ExcelConnection("tests/data/sample.xlsx") as conn:
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM Sheet1 ORDER BY id LIMIT 3")
    print(cur.fetchall())
```

## 3) Safe write pattern

```python
with ExcelConnection("tests/data/sample.xlsx", autocommit=False) as conn:
    cur = conn.cursor()
    cur.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)", (99, "Classroom"))
    conn.commit()
```

## Tool comparison

| Tool | SQL-like API | Preserves workbook formatting | Best fit |
| --- | --- | --- | --- |
| excel-dbapi | Yes (subset) | Mostly with openpyxl engine | Teaching DB-API patterns on `.xlsx` |
| pandas | No (DataFrame operations) | No (rewrites sheets) | Analysis pipelines |
| openpyxl | No (cell-oriented API) | Yes | Rich workbook manipulation |
| sqlite3 | Yes (full SQL) | N/A | Real relational storage |
