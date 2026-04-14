# excel-dbapi Roadmap

> Current line: 0.4.x series  
> SQL support reference: [`docs/SQL_SPEC.md`](SQL_SPEC.md)

---

## Completed

### Core DB-API and Engines

- PEP 249-compliant `ExcelConnection` and `ExcelCursor`
- Parameter binding with qmark placeholders (`?`)
- Openpyxl engine (default) and optional pandas engine
- Autocommit and snapshot-based rollback semantics
- Formula injection defense for write operations

### SQL Parser and Execution (Implemented)

- `SELECT` with projection, expressions, aliases, `WHERE`, `ORDER BY`, `LIMIT`, `OFFSET`
- `DISTINCT` (single-table queries)
- JOIN support: `INNER`, `LEFT`, `RIGHT`, `FULL OUTER`, `CROSS` (including chained JOINs)
- Aggregation: `COUNT`, `SUM`, `AVG`, `MIN`, `MAX`, `COUNT(DISTINCT ...)`
- `GROUP BY` and `HAVING`
- Subqueries in `WHERE ... [NOT] IN (SELECT ...)` (single-column, non-correlated)
- Compound queries: `UNION`, `UNION ALL`, `INTERSECT`, `EXCEPT`
- Expressions: arithmetic (`+ - * /`) and `CASE ... WHEN ... THEN ... ELSE ... END`
- DML: `INSERT` (single, multi-row, `INSERT ... SELECT`), `UPDATE`, `DELETE`
- UPSERT: `INSERT ... ON CONFLICT ... DO NOTHING / DO UPDATE`
- DDL: `CREATE TABLE`, `DROP TABLE`, `ALTER TABLE ADD/DROP/RENAME COLUMN`

### Quality and Delivery

- Strict mypy configuration and CI validation
- Broad parser/executor regression coverage
- Trusted Publisher (OIDC) release flow for PyPI

---

## Planned

- Async-friendly API surface (design investigation)
- Additional backend options (e.g., Polars-based engine)
- Performance tuning for large-sheet scans and join-heavy queries
- Documentation and example expansion around advanced SQL patterns
- Identifier grammar expansion (quoted identifiers and non-ASCII table/column references)

---

## Explicitly Out of Scope

- Full ACID transaction guarantees on `.xlsx`
- Multi-writer concurrency on the same workbook path
- Stored procedures, triggers, and foreign key enforcement
- Full SQL dialect parity with SQLite/PostgreSQL

---

## Documentation Authority

- Authoritative SQL behavior: [`docs/SQL_SPEC.md`](SQL_SPEC.md)
- Authoritative feature matrix: [`docs/SQL_SPEC.md#2-authoritative-feature-matrix`](SQL_SPEC.md#2-authoritative-feature-matrix)
