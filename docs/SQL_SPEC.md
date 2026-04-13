# excel-dbapi SQL Specification

> Version: 0.4  
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
| `SELECT`  | ✅ Single-table with DISTINCT / WHERE / GROUP BY / HAVING / ORDER BY / LIMIT / OFFSET / Aggregates; INNER/LEFT/RIGHT JOIN with chained JOIN clauses |
| `INSERT`  | ✅ Single-row and multi-row VALUES with optional column list; INSERT...SELECT |
| `UPDATE`  | ✅ With SET assignments and optional WHERE |
| `DELETE`  | ✅ With optional WHERE |
| `CREATE TABLE` | ✅ Creates a new worksheet with headers |
| `DROP TABLE`   | ✅ Removes a worksheet |

### 1.2 Non-Goals (Explicitly Unsupported)

The following SQL features are **rejected at parse time** with `ValueError`:

- `FULL OUTER JOIN`, `CROSS JOIN`, `NATURAL JOIN` (INNER, LEFT, and RIGHT JOIN are supported)
- `SELECT *` with JOIN (columns must be explicitly listed with table qualifiers)
- `GROUP BY`, `HAVING`, aggregates in JOIN queries
- Subqueries except `WHERE col IN (SELECT single_col FROM table [WHERE ...])` and `INSERT INTO ... SELECT ...`
- Common Table Expressions (CTEs / `WITH`)
- Window functions (`OVER`, `PARTITION BY`)
- `ALTER TABLE`
- `CREATE INDEX` / `DROP INDEX`
- `ALTER TABLE`
- `CREATE INDEX` / `DROP INDEX`
- `RETURNING`
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
  [ORDER BY column [ASC|DESC]]
  [LIMIT n]
  [OFFSET n]
```

#### JOIN (one or more tables)

```
SELECT qualified_columns FROM table [ [AS] alias ]
  { [ INNER | LEFT [OUTER] | RIGHT [OUTER] ] JOIN table [ [AS] alias ] ON condition { AND condition } }
  [WHERE conditions]
  [ORDER BY qualified_column [ASC|DESC]]
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

- Column aliases (`AS`) are **not supported**.
- Arithmetic expressions in SELECT list are **not supported** (no `col + 1`, no `UPPER(name)`).

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

### 3.3 WHERE Clause

See [Section 7: WHERE Clause](#7-where-clause).

### 3.4 ORDER BY

| Feature | Supported | Example |
|---------|-----------|---------|
| Single column | ✅ | `ORDER BY name ASC` |
| Direction | ✅ | `ASC` (default) or `DESC` |
| Multi-column | ❌ | `ORDER BY name, age` — not supported |
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
| FULL OUTER JOIN | ❌ | Not supported |
| CROSS JOIN | ❌ | Not supported |
| NATURAL JOIN | ❌ | Not supported |
| Multiple JOINs | ✅ | `... JOIN t2 ... JOIN t3 ...` |

**Requirements**:
- Table aliases are recommended (e.g., `FROM users a JOIN orders b ON ...`).
- Each table reference (alias or bare table name) must be unique; duplicate refs raise `ValueError`.
- All SELECT columns must use qualified names (`a.id`, not just `id`).
- `SELECT *` is **not supported** with JOIN.
- Subqueries (`WHERE ... IN (SELECT ...)`) are **not supported** with JOIN.
- The ON clause requires at least one equality condition (e.g., `a.id = b.user_id`).
- Multiple ON conditions are joined with `AND`.
- For chained JOINs, each ON clause may reference columns from any previously joined source and the current right source.
- `WHERE`, `ORDER BY`, `LIMIT`, `OFFSET` work with JOIN queries.
- `GROUP BY`, `HAVING`, and aggregates are **not supported** in JOIN queries.

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
| `IN` subquery | `WHERE id IN (SELECT id FROM admins WHERE role = 'admin')` | Set membership from subquery |
| `BETWEEN` | `WHERE score BETWEEN 70 AND 90` | Inclusive range |
| `LIKE` | `WHERE name LIKE 'A%'` | Pattern matching |

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
- Subqueries are supported only in `SELECT ... WHERE ... IN (SELECT ...)`
- Subqueries must select exactly one column (no `SELECT *`, no multi-column SELECT)
- Subquery form supports optional inner WHERE with literal values: `IN (SELECT id FROM admins WHERE role = 'admin')`
- Subqueries in `UPDATE ... WHERE` and `DELETE ... WHERE` are not supported
- Correlated subqueries are not supported
- Parameterized subqueries (inner `?` placeholders) are not supported

#### 7.2.3 BETWEEN Clause

- Always inclusive: `BETWEEN 1 AND 10` matches 1, 5, and 10.
- Requires the `AND` keyword between bounds.

### 7.3 Logical Connectives

| Connective | Example |
|------------|---------|
| `AND` | `WHERE x = 1 AND y = 2` |
| `OR` | `WHERE x = 1 OR y = 2` |

- Mixed `AND`/`OR`: evaluated left-to-right with `AND` binding tighter than `OR`.
  - `a AND b OR c` = `(a AND b) OR c`
  - `a OR b AND c` = `a OR (b AND c)`
- `NOT` operator is **not supported**.
- Parenthesized expressions in WHERE (e.g., `WHERE (x = 1)`) are **not supported** 
  and raise `ValueError`. Exceptions: parentheses in `IN (...)` literal lists and `IN (SELECT ...)` are allowed.

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
- Parameters can appear in: WHERE values, SET values, VALUES list, LIMIT, OFFSET

### 9.3 Binding Order

For UPDATE with WHERE:
1. SET assignment values are bound first (left to right).
2. WHERE condition values are bound next (left to right).

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
| Parenthesized WHERE | `ValueError` | `Unsupported SQL grammar: parenthesized expressions` |
| Aggregate in WHERE clause | `ValueError` | `Aggregate functions are not allowed in WHERE clause; use HAVING instead` |
| Invalid HAVING column reference | `ValueError` | `HAVING column '{column}' must be a GROUP BY column or aggregate` |
| Read-only backend mutation | `NotSupportedError` | `{action} is not supported by the read-only backend` |
| Invalid LIMIT (non-integer) | `ValueError` | `LIMIT must be an integer` |
| INSERT into headless sheet | `ValueError` | `Cannot insert into sheet without headers` |
| Empty IN clause | `ValueError` | `IN clause cannot be empty` |
| Unsupported JOIN variant | `ValueError` | `Unsupported SQL syntax: {type} JOIN` |

---

## Appendix A: Grammar (Informational EBNF)

```ebnf
statement     = compound_select | select | insert | update | delete | create | drop ;

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

where_expr    = condition { ("AND" | "OR") condition } ;
condition     = column operator value
              | column "IS" [ "NOT" ] "NULL"
              | column "IN" "(" value_list ")"
              | column "IN" "(" subquery_select ")"
              | column "BETWEEN" value "AND" value
              | column "LIKE" pattern ;

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
