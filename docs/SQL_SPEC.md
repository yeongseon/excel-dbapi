# excel-dbapi SQL Specification

> Version: **SQL Spec v1.0** (excel-dbapi 0.4.x series)
> Status: **Normative** — this document defines the frozen SQL subset for `excel-dbapi`.
> All ✅ features are stable and will not change semantics within the 0.x series.
> Last updated: 2026-04-15

---

## 1. Scope

`excel-dbapi` implements a practical SQL subset for worksheet-backed workloads. It supports
single-table CRUD, DDL, set operations, joins, aggregation, upsert, parameter binding,
and selected advanced query constructs (CTEs, window functions).

The parser is strict: unsupported grammar is rejected with `ProgrammingError`.

---

## 2. Authoritative Feature Matrix

This table is the **single authoritative matrix** for SQL feature support.

| Area | Feature | Stability | Status | Notes |
|------|---------|-----------|--------|-------|
| `SELECT` | Projection (`*`, named columns) | Stable | ✅ | `*` and explicit column lists supported |
| `SELECT` | Expressions in select list | Stable | ✅ | Arithmetic (`+ - * /`), literals, CASE |
| `SELECT` | Column aliases | Stable | ✅ | `AS alias` and implicit alias supported |
| `SELECT` | `DISTINCT` | Stable | ✅ | For `DISTINCT`, `ORDER BY` columns must be in the select list |
| `SELECT` | Window functions (`OVER (...)`) | Experimental | ⚠️ | Core support for `ROW_NUMBER`, `RANK`, `DENSE_RANK`, `SUM`, `AVG`, `COUNT`, `MIN`, `MAX` |
| `SELECT` | CTEs (`WITH`) | Experimental | ⚠️ | Non-recursive CTEs only |
| `FROM` | Table aliases | Stable | ✅ | Base table and JOIN sources |
| Identifiers | Unquoted table names | Stable | ✅ | Worksheet names follow Excel naming conventions (not strict SQL identifier grammar) |
| Identifiers | Quoted table names (`"Sales 2024"`) | Stable | ✅ | Use double quotes for names with spaces/special characters |
| Identifiers | Column names (unquoted) | Stable | ✅ | ASCII and Unicode identifiers: `[A-Za-z_\u0080-\uffff][A-Za-z0-9_\u0080-\uffff]*` |
| Identifiers | Quoted column identifiers (`"Full Name"`, `"이름"`) | Stable | ✅ | Double-quoted identifiers for columns with spaces, special characters, or reserved words |
| `WHERE` | Comparison operators | Stable | ✅ | `= == != <> > >= < <=` |
| `WHERE` | Boolean logic | Stable | ✅ | `AND`, `OR`, `NOT`, nested parentheses |
| `WHERE` | `BETWEEN` / `NOT BETWEEN` | Stable | ✅ | Inclusive bounds |
| `WHERE` | `IN` / `NOT IN` (literal list) | Stable | ✅ | Empty list rejected |
| `WHERE` | `LIKE` / `NOT LIKE` | Stable | ✅ | `%` and `_` wildcards |
| `WHERE` | `ILIKE` / `NOT ILIKE` | Stable | ✅ | Case-insensitive pattern matching |
| `WHERE` | `IS NULL` / `IS NOT NULL` | Stable | ✅ | |
| `WHERE` | Subquery in `IN` / `NOT IN` | Stable | ✅ | Single-column `SELECT`; non-correlated only |
| `WHERE` | `EXISTS (SELECT ...)` | Stable | ✅ | Correlated and non-correlated supported |
| Expressions | `CAST(expr AS type)` | Stable | ✅ | Supported target types: `INTEGER`/`INT`, `REAL`/`FLOAT`/`NUMERIC`, `TEXT`, `DATE`, `DATETIME`, `BOOLEAN` |
| Expressions | Scalar functions | Stable | ✅ | `UPPER`, `LOWER`, `LENGTH`, `TRIM`, `SUBSTR`, `COALESCE`, `NULLIF`, `CONCAT`, `YEAR`, `MONTH`, `DAY`, `ABS`, `ROUND`, `REPLACE` |
| Aggregation | `FILTER (WHERE ...)` | Stable | ✅ | Per-aggregate filtering clause |
| `JOIN` | `INNER`, `LEFT`, `RIGHT` | Stable | ✅ | Chained joins supported |
| `JOIN` | `FULL OUTER` / `FULL` | Stable | ✅ | |
| `JOIN` | `CROSS JOIN` | Stable | ✅ | No `ON` clause allowed |
| `JOIN` | `NATURAL JOIN` | — | ❌ | Rejected at parse time |
| Aggregation | `COUNT`, `SUM`, `AVG`, `MIN`, `MAX` | Stable | ✅ | |
| Aggregation | `COUNT(DISTINCT col)` | Stable | ✅ | `DISTINCT` only valid with `COUNT` |
| Grouping | `GROUP BY` | Stable | ✅ | |
| Grouping | `HAVING` | Stable | ✅ | Requires `GROUP BY` |
| Sorting | `ORDER BY` | Stable | ✅ | Multi-column, `ASC`/`DESC` |
| Pagination | `LIMIT` / `OFFSET` | Stable | ✅ | Integer literal or `?` placeholder |
| Set ops | `UNION` | Stable | ✅ | |
| Set ops | `UNION ALL` | Stable | ✅ | |
| Set ops | `INTERSECT` | Stable | ✅ | |
| Set ops | `EXCEPT` | Stable | ✅ | |
| Set ops | `INTERSECT ALL` / `EXCEPT ALL` | — | ❌ | Not implemented |
| DML | `INSERT` single-row / multi-row | Stable | ✅ | `VALUES (...)`, `VALUES (...), (...)` |
| DML | `INSERT ... SELECT` | Stable | ✅ | |
| DML | UPSERT (`ON CONFLICT`) | Stable | ✅ | `DO NOTHING`, `DO UPDATE SET ...` |
| DML | `UPDATE ... SET ... [WHERE ...]` | Stable | ✅ | SET supports row-level expressions (columns, arithmetic, functions, CAST, CASE) |
| DML | `DELETE FROM ... [WHERE ...]` | Stable | ✅ | |
| DDL | `CREATE TABLE`, `DROP TABLE` | Stable | ✅ | |
| DDL | `ALTER TABLE ADD/DROP/RENAME COLUMN` | Stable | ✅ | `COLUMN` keyword required |
| Parameters | Positional placeholders (`?`) | Stable | ✅ | qmark paramstyle |

### 2.1 Important JOIN Restrictions

- `SELECT *` in JOIN is supported, but mixing `SELECT *, col` is rejected.
- `GROUP BY` and aggregate arguments in JOIN queries must use qualified columns (`t.col`).
- Subqueries in JOIN `WHERE ... IN (SELECT ...)` are supported.

### 2.2 Identifier Grammar

Identifiers (table and column names) follow these rules:

- **Unquoted identifiers** match the pattern `[A-Za-z_\u0080-\uffff][A-Za-z0-9_\u0080-\uffff]*`.
  This includes ASCII, Korean (한글), CJK (漢字), and other Unicode letters.
- **Quoted identifiers** use SQL-standard double quotes: `"Full Name"`, `"이름"`, `"col-1"`.
  To include a literal double-quote, double it: `"col""name"` → `col"name`.
- **Qualified identifiers** use dot notation: `Sheet1.이름`, `Sheet1."Full Name"`, `"My Sheet"."My Col"`.
- Identifiers are resolved **case-insensitively** (via `casefold()`).
- Single-quoted strings (`'value'`) are always **string literals**, never identifiers.
---

## 3. Statement Syntax Overview

### Backend consistency note

- For non-transactional backends (Graph), metadata synchronization after DDL is
  best-effort. If sheet mutation succeeds but metadata sync fails, the data change
  remains applied and a warning is emitted.

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
- Set/range/pattern: `IN`, `NOT IN`, `BETWEEN`, `NOT BETWEEN`, `LIKE`, `NOT LIKE`, `ILIKE`, `NOT ILIKE`
- Logic: `AND`, `OR`, unary `NOT`, nested parentheses
- **NULL semantics**: All comparisons follow SQL three-valued logic (TRUE / FALSE / UNKNOWN). Any comparison with NULL yields UNKNOWN, which is treated as FALSE in WHERE/HAVING/ON. Use `IS NULL` / `IS NOT NULL` for explicit NULL checks.
- CASE expressions as operands

Subquery form:

```sql
WHERE col [NOT] IN (SELECT one_column FROM table [WHERE ...])
```

Subquery limitations:

- Must return exactly one column.
- No `JOIN`, `GROUP BY`, `HAVING`, `ORDER BY`, `LIMIT`, `OFFSET` in the subquery.
- Correlated subqueries: `EXISTS` only (`IN`/scalar: non-correlated only).
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

- Recursive CTEs (`WITH RECURSIVE`)
- `NATURAL JOIN`
- `SELECT ... FOR UPDATE`
- `RETURNING`
- `CREATE INDEX` / `DROP INDEX`
- `INTERSECT ALL` / `EXCEPT ALL`

---

## 8. Notes on Parser Authority

The parser implementation in `src/excel_dbapi/parser/` package is the runtime source of truth.
This specification is maintained to match implemented behavior and is updated with parser changes.

---

## 9. Feature Addition Process

New SQL features follow this lifecycle:

1. **Proposal** — Open a GitHub issue with `feat:` prefix describing the SQL construct
2. **Review** — Evaluate parser complexity, regression risk, and user need
3. **Implement** — Add to parser and executor with full test coverage
4. **Document** — Update this spec with ⚠️ Experimental stability status
5. **Stabilize** — After one minor release without semantic changes, promote to ✅ Stable

Experimental features may change semantics or be removed. Stable features will not
change semantics within the 0.x series.
