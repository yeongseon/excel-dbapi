# excel-dbapi SQL Specification

> Version: 0.6.1
> Status: **Normative** - this document defines the SQL subset supported by `excel-dbapi`.
> Last updated: 2026-04-14

---

## 1. Scope

`excel-dbapi` implements a practical SQL subset for worksheet-backed workloads. It supports
single-table CRUD, DDL, set operations, joins, aggregation, upsert, and parameter binding.

The parser is strict: unsupported grammar is rejected with `ValueError`.

---

## 2. Authoritative Feature Matrix

This table is the **single authoritative matrix** for SQL feature support.

| Area | Feature | Status | Notes |
|------|---------|--------|-------|
| `SELECT` | Projection (`*`, named columns) | ✅ | `*` and explicit column lists supported |
| `SELECT` | Expressions in select list | ✅ | Arithmetic (`+ - * /`), literals, CASE |
| `SELECT` | Column aliases | ✅ | `AS alias` and implicit alias supported |
| `SELECT` | `DISTINCT` | ✅ | Single-table queries only |
| `FROM` | Table aliases | ✅ | Base table and JOIN sources |
| `WHERE` | Comparison operators | ✅ | `= == != <> > >= < <=` |
| `WHERE` | Boolean logic | ✅ | `AND`, `OR`, `NOT`, nested parentheses |
| `WHERE` | `BETWEEN` / `NOT BETWEEN` | ✅ | Inclusive bounds |
| `WHERE` | `IN` / `NOT IN` (literal list) | ✅ | Empty list rejected |
| `WHERE` | `LIKE` / `NOT LIKE` | ✅ | `%` and `_` wildcards |
| `WHERE` | `IS NULL` / `IS NOT NULL` | ✅ | |
| `WHERE` | Subquery in `IN` / `NOT IN` | ✅ | Single-column `SELECT`; no correlated/parameterized subquery |
| `JOIN` | `INNER`, `LEFT`, `RIGHT` | ✅ | Chained joins supported |
| `JOIN` | `FULL OUTER` / `FULL` | ✅ | |
| `JOIN` | `CROSS JOIN` | ✅ | No `ON` clause allowed |
| `JOIN` | `NATURAL JOIN` | ❌ | Rejected at parse time |
| Aggregation | `COUNT`, `SUM`, `AVG`, `MIN`, `MAX` | ✅ | |
| Aggregation | `COUNT(DISTINCT col)` | ✅ | `DISTINCT` only valid with `COUNT` |
| Grouping | `GROUP BY` | ✅ | |
| Grouping | `HAVING` | ✅ | Requires `GROUP BY` |
| Sorting | `ORDER BY` | ✅ | Multi-column, `ASC`/`DESC` |
| Pagination | `LIMIT` / `OFFSET` | ✅ | Integer literal or `?` placeholder |
| Set ops | `UNION` | ✅ | |
| Set ops | `UNION ALL` | ✅ | |
| Set ops | `INTERSECT` | ✅ | |
| Set ops | `EXCEPT` | ✅ | |
| Set ops | `INTERSECT ALL` / `EXCEPT ALL` | ❌ | Not implemented |
| DML | `INSERT` single-row / multi-row | ✅ | `VALUES (...)`, `VALUES (...), (...)` |
| DML | `INSERT ... SELECT` | ✅ | |
| DML | UPSERT (`ON CONFLICT`) | ✅ | `DO NOTHING`, `DO UPDATE SET ...` |
| DML | `UPDATE ... SET ... [WHERE ...]` | ✅ | CASE in SET is supported |
| DML | `DELETE FROM ... [WHERE ...]` | ✅ | |
| DDL | `CREATE TABLE`, `DROP TABLE` | ✅ | |
| DDL | `ALTER TABLE ADD/DROP/RENAME COLUMN` | ✅ | `COLUMN` keyword required |
| Parameters | Positional placeholders (`?`) | ✅ | qmark paramstyle |

### 2.1 Important JOIN Restrictions

- `DISTINCT` with JOIN is not supported.
- `SELECT *` in JOIN is supported, but mixing `SELECT *, col` is rejected.
- `GROUP BY` and aggregate arguments in JOIN queries must use qualified columns (`t.col`).
- Subqueries in JOIN `WHERE ... IN (SELECT ...)` are not supported.

---

## 3. Statement Syntax Overview

### 3.1 SELECT

```sql
SELECT [DISTINCT] select_item [, ...]
FROM table [ [AS] alias ]
[ { [INNER|LEFT [OUTER]|RIGHT [OUTER]|FULL [OUTER]|CROSS] JOIN table [ [AS] alias ] [ON join_condition] } ... ]
[WHERE condition]
[GROUP BY column [, ...]]
[HAVING condition]
[ORDER BY order_item [, ...]]
[LIMIT n]
[OFFSET n]
```

`select_item` supports column refs, literals, arithmetic expressions, aggregate expressions,
`CASE` expressions, and aliases.

### 3.2 INSERT

```sql
INSERT INTO table [(columns)] VALUES (values)
INSERT INTO table [(columns)] VALUES (v1, ...), (v2, ...)
INSERT INTO table [(columns)] SELECT ...
INSERT INTO table [(columns)] ... ON CONFLICT (col [, ...]) DO NOTHING
INSERT INTO table [(columns)] ... ON CONFLICT (col [, ...]) DO UPDATE SET col = expr [, ...]
```

### 3.3 UPDATE

```sql
UPDATE table SET assignment [, ...] [WHERE condition]
```

### 3.4 DELETE

```sql
DELETE FROM table [WHERE condition]
```

### 3.5 DDL

```sql
CREATE TABLE table (col [TYPE], ...)
DROP TABLE table
ALTER TABLE table ADD COLUMN col TYPE
ALTER TABLE table DROP COLUMN col
ALTER TABLE table RENAME COLUMN old TO new
```

### 3.6 Compound Queries

```sql
SELECT ...
UNION [ALL] | INTERSECT | EXCEPT
SELECT ...
```

Compound branches are evaluated left-to-right.

---

## 4. WHERE Clause Semantics

Supported:

- Comparisons: `=`, `==`, `!=`, `<>`, `>`, `>=`, `<`, `<=`
- Null checks: `IS NULL`, `IS NOT NULL`
- Set/range/pattern: `IN`, `NOT IN`, `BETWEEN`, `NOT BETWEEN`, `LIKE`, `NOT LIKE`
- Logic: `AND`, `OR`, unary `NOT`, nested parentheses
- CASE expressions as operands

Subquery form:

```sql
WHERE col [NOT] IN (SELECT one_column FROM table [WHERE ...])
```

Subquery limitations:

- Must return exactly one column.
- No `JOIN`, `GROUP BY`, `HAVING`, `ORDER BY`, `LIMIT`, `OFFSET` in the subquery.
- No correlated references.
- No placeholders inside the subquery.

---

## 5. Aggregates, GROUP BY, HAVING

- Aggregates: `COUNT(*)`, `COUNT(col)`, `COUNT(DISTINCT col)`, `SUM`, `AVG`, `MIN`, `MAX`
- `HAVING` requires `GROUP BY`.
- In JOIN queries, aggregate args (except `COUNT(*)`) must be qualified (`t.col`).
- Non-aggregate selected columns must be compatible with grouping semantics.

---

## 6. Parameter Binding

- Paramstyle: **qmark** (`?`, PEP 249)
- Binding order for SELECT-family statements:
  1. SELECT expressions (including CASE)
  2. WHERE
  3. HAVING
  4. ORDER BY expressions
  5. LIMIT/OFFSET
- UPDATE binds SET values before WHERE values.
- INSERT multi-row and UPSERT placeholders are bound left-to-right in SQL order.

---

## 7. Explicitly Unsupported Features

- CTEs (`WITH`)
- Window functions (`OVER`, `PARTITION BY`)
- `NATURAL JOIN`
- `SELECT ... FOR UPDATE`
- `RETURNING`
- `CREATE INDEX` / `DROP INDEX`
- `INTERSECT ALL` / `EXCEPT ALL`

---

## 8. Notes on Parser Authority

The parser implementation in `src/excel_dbapi/parser.py` is the runtime source of truth.
This specification is maintained to match implemented behavior and is updated with parser changes.
