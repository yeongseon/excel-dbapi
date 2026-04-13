# excel-dbapi SQL Specification

> Version: 0.6.1
> Status: **Normative** — this document defines the SQL subset that excel-dbapi supports.  
> Last updated: 2026-04-13

---

## 1. Scope

excel-dbapi implements a **minimal SQL subset** designed for single-table CRUD operations
against Excel worksheets. It is intentionally limited — complex queries should use a real
database engine.

### 1.1 Supported Statements

| Statement | Supported |
|-----------|-----------|
| `SELECT`  | ✅ Single-table with DISTINCT / WHERE / GROUP BY / HAVING / multi-column ORDER BY / LIMIT / OFFSET / Aggregates; INNER/LEFT/RIGHT JOIN with chained JOIN clauses |
| `INSERT`  | ✅ Single-row and multi-row VALUES with optional column list; INSERT...SELECT |
| `UPDATE`  | ✅ With SET assignments and optional WHERE |
| `DELETE`  | ✅ With optional WHERE |
| `CREATE TABLE` | ✅ Creates a new worksheet with headers |
| `DROP TABLE`   | ✅ Removes a worksheet |
| `ALTER TABLE`  | ✅ `ADD COLUMN`, `DROP COLUMN`, `RENAME COLUMN` |

### 1.2 Non-Goals (Explicitly Unsupported)

The following SQL features are **rejected at parse time** with `ValueError`:

- `NATURAL JOIN` (INNER, LEFT, RIGHT, FULL OUTER, and CROSS JOIN are supported)
- Mixed `SELECT *, col` with JOIN (bare `SELECT *` is supported; see §4.8)
- `GROUP BY` / aggregate arguments in JOIN queries that use bare (unqualified) columns (for example `GROUP BY dept`, `SUM(amount)`)
- Subqueries except `WHERE col [NOT] IN (SELECT single_col FROM table [WHERE ...])` and `INSERT INTO ... SELECT ...`
- Common Table Expressions (CTEs / `WITH`)
- Window functions (`OVER`, `PARTITION BY`)
- `CREATE INDEX` / `DROP INDEX`
- `RETURNING`
- `SELECT ... FOR UPDATE`

> **Design principle**: It is better to reject unsupported SQL loudly than to silently
> produce wrong results by ignoring clauses.

---

## 2. Lexical Rules

### 2.1 Identifiers

- Column and table names are **unquoted, case-sensitive** tokens.
- Table names correspond to Excel worksheet names.
- Column names correspond to the first row (header row) of each worksheet.
- Reserved words (`SELECT`, `FROM`, `WHERE`, etc.) are **case-insensitive**.

#### 2.1.1 Header Validation

When a worksheet is loaded, the header row is validated and normalized:

- **Type coercion**: Non-string header values (e.g., numeric `1`, `2.5`) are coerced
  to their string representation (`"1"`, `"2.5"`).
- **Empty headers**: `None`, empty string (`""`), or whitespace-only headers raise
  `DataError` — every column must have a meaningful name.
- **Duplicate detection**: Headers are checked for uniqueness in a **case-insensitive**
  manner. For example, `["Name", "name"]` raises `DataError` because they collide
  after lowercasing. The original casing is preserved for non-duplicate headers.
- **Whitespace trimming**: Leading and trailing whitespace is stripped from each header
  before validation.

> **Rationale**: Excel allows arbitrary cell values in the header row, but SQL column
> names must be non-empty and unique. Validating at load time prevents silent data
> corruption caused by `dict(zip(headers, row))` when duplicate headers exist.

### 2.2 String Literals

- Single-quoted: `'hello world'`
- Escaped single quotes use doubling: `'it''s'` → `it's`
- Double-quoted: `"value"` (also supported)
- Escaped double quotes: `"say ""hello"""` → `say "hello"`
- Strings may contain spaces, preserved as-is.

### 2.3 Numeric Literals

- Integers: `42`, `-1`, `0`
- Floats: `3.14`, `-0.5`
- Parsed via Python's `int()` then `float()` fallback.

### 2.4 NULL Literal

- `NULL` (case-insensitive) → Python `None`

### 2.5 Placeholders

- Positional placeholders: `?` (qmark paramstyle, PEP 249)
- Bound at parse time via the `params` tuple.

### 2.6 Tokenization

The SQL string is tokenized with the following rules:
1. Whitespace separates tokens.
2. Quoted strings (single or double) are preserved as single tokens, including internal spaces.
3. Parentheses `(` and `)` are emitted as standalone tokens.
4. No support for block comments (`/* */`) or line comments (`--`).

---

## 3. SELECT

### 3.1 Syntax

#### Single-table

```
SELECT [DISTINCT] columns FROM table
  [WHERE conditions]
  [GROUP BY column { "," column }]
  [HAVING conditions]
  [ORDER BY column [ASC|DESC] { "," column [ASC|DESC] }]
  [LIMIT n]
  [OFFSET n]
```

#### JOIN (one or more tables)

```
SELECT qualified_columns FROM table [ [AS] alias ]
  { [ INNER | LEFT [OUTER] | RIGHT [OUTER] ] JOIN table [ [AS] alias ] ON condition { AND condition } }
  [WHERE conditions]
  [GROUP BY qualified_column { "," qualified_column }]
  [HAVING conditions]
  [ORDER BY qualified_column [ASC|DESC] { "," qualified_column [ASC|DESC] }]
  [LIMIT n]
  [OFFSET n]
```

### 3.2 Columns

| Form | Example | Description |
|------|---------|-------------|
| Wildcard | `SELECT * FROM Sheet1` | All columns in header order |
| Named | `SELECT id, name FROM Sheet1` | Specific columns (must exist in header) |
| Aggregate | `SELECT COUNT(*) FROM Sheet1` | Aggregate function call |
| Mixed | `SELECT name, COUNT(*) FROM Sheet1 GROUP BY name` | Plain + aggregate columns |
| Qualified | `SELECT a.id, b.name FROM Sheet1 a JOIN Sheet2 b ON a.id = b.id` | Table-qualified columns (required in JOIN) |
| Arithmetic expression | `SELECT price * qty AS total FROM Sheet1` | Row-level arithmetic with `+`, `-`, `*`, `/`, unary `-`, and parentheses |
| CASE expression | `SELECT CASE WHEN score > 10 THEN 'high' ELSE 'low' END AS band FROM Sheet1` | Row-level conditional expression |

- Column aliases (`AS`) are **not supported**.
- Function expressions in SELECT list are **not supported** (for example `UPPER(name)`).

#### 3.2.1 Aggregate Functions

| Function | Example | Description |
|----------|---------|-------------|
| `COUNT(*)` | `SELECT COUNT(*) FROM Sheet1` | Count all rows |
| `COUNT(col)` | `SELECT COUNT(score) FROM Sheet1` | Count non-NULL values |
| `SUM(col)` | `SELECT SUM(score) FROM Sheet1` | Sum of numeric values |
| `AVG(col)` | `SELECT AVG(score) FROM Sheet1` | Average of numeric values |
| `MIN(col)` | `SELECT MIN(score) FROM Sheet1` | Minimum numeric value |
| `MAX(col)` | `SELECT MAX(score) FROM Sheet1` | Maximum numeric value |

**NULL handling**:
- `COUNT(*)` on empty set → `0`
- `COUNT(col)` excludes NULL values
- `SUM/AVG/MIN/MAX` on empty set or all-NULL → `None`
- Non-numeric values are ignored by `SUM/AVG/MIN/MAX`

**Implicit grouping**: Without `GROUP BY`, aggregates treat the entire result set as one group and return a single row.

**Mixed columns**: When non-aggregate columns appear alongside aggregates, `GROUP BY` is required. Non-aggregate columns must appear in `GROUP BY`.

**Aggregate arguments**: Aggregate calls accept only a bare column name (e.g., `SUM(score)`) or `*` for `COUNT(*)`. Expressions such as `COUNT(DISTINCT name)` and `SUM(score + 1)` are not supported.

In JOIN queries, aggregate arguments must use qualified column names (for example `SUM(t2.amount)`), except `COUNT(*)`.

#### 3.2.2 Arithmetic Expressions

Arithmetic expressions are evaluated per-row in the `SELECT` list.

**Supported operators**:
- Binary: `+`, `-`, `*`, `/`
- Unary: leading `-` (negation)
- Parentheses for grouping: `(a + b) * 2`

**Precedence and associativity**:
1. Unary `-`
2. `*`, `/`
3. `+`, `-`
4. Left-associative for binary operators

**Operands**:
- Bare column names (single-table)
- Qualified columns (`source.column`) in JOIN queries
- Numeric literals (integer and float)

**NULL propagation**:
- If any operand in a binary arithmetic operation is `NULL`, the result is `NULL`.
- Unary negation of `NULL` returns `NULL`.

**Error behavior**:
- Division by zero raises `ProgrammingError`.
- Non-numeric operands raise `ProgrammingError`.

**Current limitations**:
- Arithmetic expressions are supported only in the `SELECT` list (not in `WHERE`, `GROUP BY`, `HAVING`, or `ORDER BY` expressions).
- Aggregate arguments still accept only bare column names or `*` (for example, `SUM(a * b)` is rejected).

#### 3.2.3 CASE Expressions

`CASE` expressions are supported in the SELECT list and may be nested.

**Searched CASE**:
- `CASE WHEN condition THEN result [WHEN condition THEN result ...] [ELSE result] END`

**Simple CASE**:
- `CASE expr WHEN match THEN result [WHEN match THEN result ...] [ELSE result] END`

**Rules**:
- `WHEN` conditions in searched CASE use standard WHERE-condition semantics.
- `THEN`, `ELSE`, and simple `WHEN match` values accept scalar expressions (column refs, literals, arithmetic, nested CASE).
- If no `WHEN` branch matches and `ELSE` is omitted, the result is `NULL`.
- Parameter placeholders (`?`) are supported in `WHEN` conditions and `THEN`/`ELSE` results.

### 3.3 WHERE Clause

See [Section 7: WHERE Clause](#7-where-clause).

### 3.4 ORDER BY

| Feature | Supported | Example |
|---------|-----------|---------|
| Single column | ✅ | `ORDER BY name ASC` |
| Direction | ✅ | `ASC` (default) or `DESC` |
| Multi-column | ✅ | `ORDER BY name ASC, age DESC` |
| Expressions | ❌ | `ORDER BY UPPER(name)` — not supported |

- Non-aggregate queries: ORDER BY column must exist in worksheet headers.
- Aggregate queries (`GROUP BY` and/or aggregate SELECT): ORDER BY may reference:
  - projected output columns,
  - GROUP BY columns (even if not projected),
  - aggregate expressions (e.g., `COUNT(*)`, `SUM(score)`).
- NULL values sort last (after all non-NULL values).
- Numeric strings are compared numerically when both operands parse as numbers.

### 3.5 LIMIT

| Feature | Supported | Example |
|---------|-----------|---------|
| Integer literal | ✅ | `LIMIT 10` |
| Placeholder | ✅ | `LIMIT ?` |
| With OFFSET | ✅ | `LIMIT 10 OFFSET 5` |

### 3.6 OFFSET

| Feature | Supported | Example |
|---------|-----------|---------|
| Integer literal | ✅ | `OFFSET 5` |
| Placeholder | ✅ | `OFFSET ?` |
| Bare OFFSET (without LIMIT) | ✅ | `SELECT * FROM t OFFSET 5` |

### 3.7 DISTINCT

- `DISTINCT` is supported immediately after `SELECT`.
- Deduplication is applied after projection and preserves first-seen row order.

### 3.8 GROUP BY

| Feature | Supported | Example |
|---------|-----------|---------|
| Single column | ✅ | `GROUP BY name` |
| Multiple columns | ✅ | `GROUP BY dept, name` |

- Partitions rows into groups by the specified column(s).
- Columns in the SELECT list must either be aggregate functions or appear in GROUP BY.
- Groups preserve insertion order (dict key order).

### 3.9 HAVING

| Feature | Supported | Example |
|---------|-----------|---------|
| Aggregate condition | ✅ | `HAVING COUNT(*) > 1` |
| Comparison | ✅ | `HAVING SUM(score) >= 100` |
| Multiple conditions | ✅ | `HAVING COUNT(*) > 1 AND SUM(score) > 50` |

- Requires `GROUP BY` — `HAVING` without `GROUP BY` raises `ValueError`.
- Filters groups after aggregation (contrast with `WHERE` which filters rows before grouping).
- Conditions may reference aggregate expressions (e.g., `SUM(score)`) and `GROUP BY` columns.

### 3.10 Clause Ordering

Clauses must appear in this order: `WHERE` → `GROUP BY` → `HAVING` → `ORDER BY` → `LIMIT` → `OFFSET`.
Any other ordering raises `ValueError`.

### 3.11 JOIN

| Feature | Supported | Example |
|---------|-----------|---------|
| INNER JOIN | ✅ | `SELECT a.id, b.name FROM t1 a JOIN t2 b ON a.id = b.id` |
| LEFT JOIN | ✅ | `SELECT a.id, b.name FROM t1 a LEFT JOIN t2 b ON a.id = b.id` |
| RIGHT JOIN | ✅ | `SELECT a.id, b.name FROM t1 a RIGHT JOIN t2 b ON a.id = b.id` |
| FULL OUTER JOIN | ✅ | `A FULL OUTER JOIN B ON ...` (or `FULL JOIN`) |
| CROSS JOIN | ✅ | `A CROSS JOIN B` (no ON clause; cartesian product) |
| NATURAL JOIN | ❌ | Not supported |
| Multiple JOINs | ✅ | `... JOIN t2 ... JOIN t3 ...` |

**Requirements**:
- Table aliases are recommended (e.g., `FROM users a JOIN orders b ON ...`).
- Each table reference (alias or bare table name) must be unique; duplicate refs raise `ValueError`.
- All SELECT columns must use qualified names (`a.id`, not just `id`).
- `SELECT *` is supported with JOIN and expands to all columns from all joined tables in JOIN clause order (left table first, then each joined table), using qualified names (`source.column`).
- Subqueries (`WHERE ... IN (SELECT ...)`) are **not supported** with JOIN.
- The ON clause requires at least one equality condition (e.g., `a.id = b.user_id`).
- Multiple ON conditions are joined with `AND`.
- `CROSS JOIN` produces a cartesian product of all rows; `ON` clause is rejected for CROSS JOIN.
- `FULL OUTER JOIN` preserves all rows from both sides, filling NULLs where no match exists.
- For chained JOINs, each ON clause may reference columns from any previously joined source and the current right source.
- `WHERE`, `ORDER BY`, `LIMIT`, `OFFSET` work with JOIN queries.
- `GROUP BY`, `HAVING`, and aggregates are supported in JOIN queries.
- JOIN + GROUP BY requires qualified column names in `GROUP BY` (`GROUP BY t1.dept`, not `GROUP BY dept`).
- JOIN aggregates require qualified arguments (for example `SUM(t2.amount)`), except `COUNT(*)`.

**Execution**:
- INNER JOIN returns only rows where the ON condition matches in both tables.
- LEFT JOIN returns all rows from the left table, with NULL values for unmatched right-table columns.
- RIGHT JOIN returns all rows from the right table, with NULL values for unmatched left-table columns.
- Chained JOINs execute left-to-right as iterative two-source folds.
- The join algorithm uses hash matching for efficient lookups.

### 3.12 Compound Queries (Set Operations)

| Operation | Supported | Behavior |
|-----------|-----------|----------|
| `UNION` | ✅ | Combines rows and removes duplicates |
| `UNION ALL` | ✅ | Combines rows and keeps duplicates |
| `INTERSECT` | ✅ | Returns rows present in all inputs |
| `EXCEPT` | ✅ | Returns rows in the left result that are absent from the right result |

**Syntax**:

```
SELECT ... FROM ...
  (UNION [ALL] | INTERSECT | EXCEPT)
SELECT ... FROM ...
  [ (UNION [ALL] | INTERSECT | EXCEPT) SELECT ... ] ...
```

**Rules**:
- Each branch must be a valid `SELECT` query.
- All branches must return the same number of columns.
- Chained compounds are evaluated left-to-right: `(A UNION B) EXCEPT C`.
- Branch-local clauses (`WHERE`, `ORDER BY`, `LIMIT`, `OFFSET`, etc.) are allowed within each branch.

**Examples**:
- `SELECT id, name FROM users UNION SELECT id, name FROM admins`
- `SELECT id FROM t1 UNION ALL SELECT id FROM t2`
- `SELECT id FROM t1 INTERSECT SELECT id FROM t2`
- `SELECT id FROM t1 EXCEPT SELECT id FROM t2`
- `SELECT id FROM a UNION SELECT id FROM b UNION SELECT id FROM c`

---

## 4. INSERT

### 4.1 Syntax

```
INSERT INTO table [(columns)] VALUES (values)
INSERT INTO table [(columns)] VALUES (v1, v2), (v3, v4), ...
INSERT INTO table [(columns)] SELECT columns FROM source [WHERE ...]
```

### 4.2 Forms

| Form | Example |
|------|---------|
| With columns | `INSERT INTO Sheet1 (id, name) VALUES (1, 'Alice')` |
| Without columns | `INSERT INTO Sheet1 VALUES (1, 'Alice')` |
| With placeholders | `INSERT INTO Sheet1 (id, name) VALUES (?, ?)` |

### 4.3 Rules

- Column count must match value count.
- If columns are omitted, value count must match header count.
- Columns must exist in the worksheet header.
- Multi-row INSERT is supported: each value tuple is inserted as a separate row.
- INSERT...SELECT is supported: rows from a SELECT query are inserted into the target table.
- For INSERT...SELECT, column count of the SELECT result must match the target column count.
- If SELECT returns zero rows, no rows are inserted and `rowcount` is 0.
- Parameter binding in multi-row VALUES works across all tuples: `VALUES (?, ?), (?, ?)` with 4 params.
- Formula injection defense: values starting with `=`, `+`, `-`, `@`, `\t`, `\r`
  are prefixed with `'` by default (configurable via `sanitize_formulas=False`).

---

## 5. UPDATE

### 5.1 Syntax

```
UPDATE table SET assignments [WHERE conditions]
```

### 5.2 Assignments

| Form | Example |
|------|---------|
| Literal | `SET name = 'Bob'` |
| Placeholder | `SET name = ?` |
| Multiple | `SET name = 'Bob', age = 30` |
| NULL | `SET name = NULL` |
| CASE expression | `SET status = CASE WHEN score >= 60 THEN 'pass' ELSE 'fail' END` |

### 5.3 Rules

- All assignment columns must exist in the worksheet header.
- Without WHERE, all rows are updated.
- Parameter binding processes SET values before WHERE values.

---

## 6. DELETE

### 6.1 Syntax

```
DELETE FROM table [WHERE conditions]
```

### 6.2 Rules

- Without WHERE, all data rows are deleted (headers preserved).
- With WHERE, only matching rows are removed.

---

## 7. WHERE Clause

### 7.1 Comparison Operators

| Operator | Example | Description |
|----------|---------|-------------|
| `=`, `==` | `WHERE id = 1` | Equality |
| `!=`, `<>` | `WHERE id != 1` | Inequality |
| `>` | `WHERE score > 80` | Greater than |
| `>=` | `WHERE score >= 80` | Greater than or equal |
| `<` | `WHERE score < 80` | Less than |
| `<=` | `WHERE score <= 80` | Less than or equal |

### 7.2 Special Operators

| Operator | Example | Description |
|----------|---------|-------------|
| `IS NULL` | `WHERE name IS NULL` | NULL check |
| `IS NOT NULL` | `WHERE name IS NOT NULL` | Non-NULL check |
| `IN` | `WHERE name IN ('Alice', 'Bob')` | Set membership |
| `NOT IN` | `WHERE name NOT IN ('Alice', 'Bob')` | Negated set membership |
| `IN` subquery | `WHERE id IN (SELECT id FROM admins WHERE role = 'admin')` | Set membership from subquery |
| `NOT IN` subquery | `WHERE id NOT IN (SELECT id FROM admins)` | Negated set membership from subquery |
| `BETWEEN` | `WHERE score BETWEEN 70 AND 90` | Inclusive range |
| `NOT BETWEEN` | `WHERE score NOT BETWEEN 70 AND 90` | Negated inclusive range |
| `LIKE` | `WHERE name LIKE 'A%'` | Pattern matching |
| `NOT LIKE` | `WHERE name NOT LIKE 'A%'` | Negated pattern matching |

#### 7.2.1 LIKE Patterns

| Pattern | Matches |
|---------|---------|
| `%` | Any sequence of characters (including empty) |
| `_` | Any single character |

Example: `WHERE name LIKE 'A%'` matches "Alice", "Ann", "A".

#### 7.2.2 IN Clause

- Values must be parenthesized: `IN (1, 2, 3)` or `IN ('a', 'b')`
- Empty IN list raises `ValueError`
- Supports placeholder binding: `IN (?, ?, ?)`
- Subqueries are supported in `SELECT`, `UPDATE`, and `DELETE` WHERE clauses: `WHERE col [NOT] IN (SELECT ...)`
- Subqueries must select exactly one column (no `SELECT *`, no multi-column SELECT)
- Subquery form supports optional inner WHERE with literal values: `IN (SELECT id FROM admins WHERE role = 'admin')`
- Correlated subqueries are not supported
- Parameterized subqueries (inner `?` placeholders) are not supported

#### 7.2.3 BETWEEN Clause

- Always inclusive: `BETWEEN 1 AND 10` matches 1, 5, and 10.
- Requires the `AND` keyword between bounds.

#### 7.2.4 CASE as Condition Operand

CASE expressions can appear as condition operands in WHERE clauses.

- Example: `WHERE CASE WHEN score > 10 THEN 'a' ELSE 'b' END = 'a'`
- Both searched and simple CASE forms are supported.
- Nested CASE expressions are supported.

### 7.3 Logical Connectives

| Connective | Example |
|------------|---------|
| `AND` | `WHERE x = 1 AND y = 2` |
| `OR` | `WHERE x = 1 OR y = 2` |

- Mixed `AND`/`OR`: evaluated left-to-right with `AND` binding tighter than `OR`.
  - `a AND b OR c` = `(a AND b) OR c`
  - `a OR b AND c` = `a OR (b AND c)`
- `NOT` operator: negates a single condition or a parenthesized group.
  - `WHERE NOT x = 1` — negates the condition
  - `WHERE NOT (x = 1 OR y = 2)` — negates the parenthesized group
  - `NOT NOT x = 1` — double negation is allowed
- Parenthesized expressions control grouping and precedence.
  - `WHERE (x = 1 OR y = 2) AND z = 3` — OR evaluated before AND
  - `WHERE (a = 1 AND b = 2) OR (c = 3 AND d = 4)` — two grouped ANDs with OR
  - Nested parentheses are supported: `WHERE ((x = 1))`

### 7.4 Type Coercion

When comparing values, the engine attempts numeric coercion:
- If both values parse as numbers (int/float), they are compared numerically.
- Otherwise, both are compared as strings via `str()`.
- `None` compared with any non-IS operator returns `False`.
- Boolean values are not coerced to numbers.

---

## 8. DDL (Data Definition Language)

### 8.1 CREATE TABLE

```
CREATE TABLE name (col1 [TYPE], col2 [TYPE], ...)
```

- Creates a new worksheet with the specified column names as headers.
- Type annotations after column names are accepted but only the column name is used.
- If the worksheet already exists, raises `ValueError`.

### 8.2 DROP TABLE

```
DROP TABLE name
```

- Removes the worksheet from the workbook.
- If the worksheet does not exist, raises `ValueError`.

### 8.3 ALTER TABLE

```
ALTER TABLE name ADD COLUMN column_name type_name
ALTER TABLE name DROP COLUMN column_name
ALTER TABLE name RENAME COLUMN old_name TO new_name
```

**Rules**:
- `COLUMN` keyword is required for all ALTER variants.
- `ADD COLUMN` appends the new column to the end of the header list.
- Existing rows receive `NULL` for newly added columns.
- `DROP COLUMN` removes the column from headers and all row values at that position.
- `RENAME COLUMN` changes only the header text; row values are preserved.

**Supported types for `ADD COLUMN`**:
- `TEXT`, `INTEGER`, `REAL`, `FLOAT`, `BOOLEAN`, `DATE`, `DATETIME`
- `FLOAT` is normalized to `REAL`.

> **Note**: The type name in `ADD COLUMN` is validated during parsing but is
> **not persisted** in the worksheet. Excel worksheets store values without
> schema types. Column types are inferred from stored values at reflection
> time. A newly added column with no values will be reflected as `TEXT`.

**Error conditions**:
- Invalid syntax or missing `COLUMN` keyword raises `ValueError`.
- Unknown `ADD COLUMN` type raises `ValueError`.
- Altering a missing table raises `ValueError`.
- Adding an existing column raises `ValueError`.
- Dropping a non-existent column raises `ValueError`.
- Dropping the only remaining column raises `ValueError`.
- Renaming from a non-existent source column raises `ValueError`.
- Renaming to an already existing target column raises `ValueError`.

---

## 9. Parameter Binding

### 9.1 Paramstyle

excel-dbapi uses **qmark** paramstyle (PEP 249):

```python
cursor.execute("SELECT * FROM Sheet1 WHERE id = ?", (42,))
cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)", (1, "Alice"))
```

### 9.2 Binding Rules

- Each `?` in the query consumes one parameter from the `params` tuple, left to right.
- Too few parameters → `ValueError: Not enough parameters for placeholders`
- Too many parameters → `ValueError: Too many parameters for placeholders`
- Parameters can appear in: SELECT CASE branches, WHERE values, SET values (including CASE branches), VALUES list, LIMIT, OFFSET

### 9.3 Binding Order

For UPDATE with WHERE:
1. SET assignment values are bound first (left to right).
2. WHERE condition values are bound next (left to right).

For SELECT with WHERE/HAVING/LIMIT/OFFSET:
1. SELECT expression placeholders (including CASE branches) are bound first, left to right.
2. WHERE values are bound next.
3. HAVING values are bound next.
4. LIMIT and OFFSET placeholders are bound last.

Within CASE expressions, placeholders bind in SQL order:
- Searched CASE: each `WHEN` condition first, then its `THEN` result, then `ELSE`.
- Simple CASE: CASE value first, then each `WHEN match`, then `THEN` result, then `ELSE`.

---

## 10. Error Behavior

| Condition | Exception | Message Pattern |
|-----------|-----------|-----------------|
| Unknown SQL action | `ValueError` | `Unsupported SQL action: {action}` |
| Missing FROM in SELECT | `ValueError` | `Invalid SQL query format: ...` |
| Table not found | `ValueError` | `Sheet '{name}' not found in Excel` |
| Column not found | `ValueError` | `Unknown column(s): ...` |
| Sheet already exists (CREATE) | `ValueError` | `Sheet '{name}' already exists` |
| Param count mismatch | `ValueError` | `Not enough / Too many parameters` |
| Unsupported grammar | `ValueError` | `Unsupported SQL grammar: {feature}` |
| Aggregate in WHERE clause | `ValueError` | `Aggregate functions are not allowed in WHERE clause; use HAVING instead` |
| Invalid HAVING column reference | `ValueError` | `HAVING column '{column}' must be a GROUP BY column or aggregate` |
| Read-only backend mutation | `NotSupportedError` | `{action} is not supported by the read-only backend` |
| Invalid LIMIT (non-integer) | `ValueError` | `LIMIT must be an integer` |
| INSERT into headless sheet | `ValueError` | `Cannot insert into sheet without headers` |
| Empty IN clause | `ValueError` | `IN clause cannot be empty` |
| Unsupported JOIN variant | `ValueError` | `Unsupported SQL syntax: {type} JOIN` |
| Invalid header (empty/None) | `DataError` | `Empty or None header at column index {idx}` |
| Duplicate header (case-insensitive) | `DataError` | `Duplicate header: {h!r} (conflicts with {existing!r})` |

---

## 11. Transactional Behavior

### 11.1 Autocommit

When `autocommit=True` (the default), each mutating statement (`INSERT`,
`UPDATE`, `DELETE`, `CREATE TABLE`, `DROP TABLE`, `ALTER TABLE`) is
automatically saved to disk after successful execution. `SELECT` and other
read-only statements do not trigger a save.

### 11.2 `executemany` Atomicity

`cursor.executemany(sql, seq_of_params)` executes the same statement for
each parameter set in `seq_of_params`. The save behavior is **atomic**:

- A snapshot of the workbook is taken before the batch begins.
- Each parameter set is applied sequentially (in-memory only).
- If **all** executions succeed, the workbook is saved **once** at the end.
- If **any** execution fails, the workbook is restored to the pre-batch
  snapshot and the exception is re-raised. No partial mutations are persisted.

> **Rationale**: Per-row saves would be both slow (O(n) disk writes) and
> semantically surprising — a mid-batch failure would leave the workbook in
> a half-mutated state. Atomic batch semantics match the behavior of
> database drivers that wrap `executemany` in an implicit transaction.

---

## Appendix A: Grammar (Informational EBNF)

```ebnf
statement     = compound_select | select | insert | update | delete | create | drop | alter ;

compound_select = select compound_op select { compound_op select } ;
compound_op   = ( "UNION" [ "ALL" ] ) | "INTERSECT" | "EXCEPT" ;

select        = "SELECT" [ "DISTINCT" ] select_columns "FROM" table_ref
                [ join_clause ]
                [ "WHERE" where_expr ]
                [ "GROUP" "BY" column { "," column } ]
                [ "HAVING" where_expr ]
                [ "ORDER" "BY" qualified_column [ direction ] ]
                [ "LIMIT" integer ]
                [ "OFFSET" integer ] ;

insert        = "INSERT" "INTO" table [ "(" column_list ")" ]
                ( "VALUES" value_tuple { "," value_tuple }
                | select ) ;

value_tuple   = "(" value_list ")" ;

update        = "UPDATE" table "SET" assignment { "," assignment }
                [ "WHERE" where_expr ] ;

delete        = "DELETE" "FROM" table [ "WHERE" where_expr ] ;

create        = "CREATE" "TABLE" table "(" column_def { "," column_def } ")" ;
drop          = "DROP" "TABLE" table ;
alter         = "ALTER" "TABLE" table alter_op ;
alter_op      = ( "ADD" "COLUMN" column type_name )
              | ( "DROP" "COLUMN" column )
              | ( "RENAME" "COLUMN" column "TO" column ) ;

select_columns = "*" | select_item { "," select_item } ;
select_item    = column | aggregate ;
aggregate      = aggregate_func "(" ( "*" | column ) ")" ;
aggregate_func = "COUNT" | "SUM" | "AVG" | "MIN" | "MAX" ;
columns       = "*" | column { "," column } ;
column_list   = column { "," column } ;
value_list    = value { "," value } ;
column_def    = column [ type_name ] ;
assignment    = column "=" value ;
qualified_col  = table_or_alias "." column ;
table_ref      = table [ [ "AS" ] alias ] ;
join_clause    = { [ "INNER" | "LEFT" [ "OUTER" ] | "RIGHT" [ "OUTER" ] ] "JOIN" table_ref "ON" join_cond { "AND" join_cond } } ;
join_cond      = qualified_col "=" qualified_col ;
alias          = identifier ;

where_expr    = or_expr ;
or_expr       = and_expr { "OR" and_expr } ;
and_expr      = not_expr { "AND" not_expr } ;
not_expr      = "NOT" not_expr | factor ;
factor        = "(" or_expr ")" | condition ;
condition     = column operator value
              | column "IS" [ "NOT" ] "NULL"
              | column [ "NOT" ] "IN" "(" value_list ")"
              | column [ "NOT" ] "IN" "(" subquery_select ")"
              | column [ "NOT" ] "BETWEEN" value "AND" value
              | column [ "NOT" ] "LIKE" pattern ;

subquery_select = "SELECT" column "FROM" table [ "WHERE" where_expr ] ;

operator      = "=" | "==" | "!=" | "<>" | ">" | ">=" | "<" | "<=" ;
direction     = "ASC" | "DESC" ;
value         = string | number | "NULL" | "?" ;
string        = "'" { character } "'" | '"' { character } '"' ;
number        = integer | float ;
integer       = digit { digit } ;
float         = digit { digit } "." digit { digit } ;
pattern       = string ;  (* containing % and _ wildcards *)
table         = identifier ;
column        = identifier ;
type_name     = identifier ;
identifier    = letter { letter | digit | "_" } ;
```

> **Note**: This EBNF is informational. The authoritative behavior is defined by the
> `excel_dbapi.parser` module and the sections above.
