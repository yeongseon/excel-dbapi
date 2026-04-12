# excel-dbapi Roadmap

> **Current version**: 0.3.0 · **Python**: 3.10+ · **Published**: [PyPI](https://pypi.org/project/excel-dbapi/)

## Completed

### v0.1.0 — PEP 249 Foundation

- PEP 249 (DB-API 2.0) compliant `ExcelConnection` and `ExcelCursor`
- SQL parser: SELECT with WHERE (AND/OR, comparison operators), ORDER BY, LIMIT
- INSERT, UPDATE, DELETE execution
- CREATE TABLE / DROP TABLE (DDL)
- Openpyxl engine (default) and Pandas engine (optional)
- Transaction simulation: `commit()` / `rollback()` with in-memory snapshot
- PEP 249 exception hierarchy
- Parameter binding (`?` placeholder)
- Formula injection defense (enabled by default)
- Reflection helpers for dialect integration
- Metadata sheet for schema persistence

### v0.2.x — Operators & Quality

- IN, BETWEEN, LIKE operators for WHERE clauses
- Codecov CI integration
- mypy strict mode enabled and passing
- PyPI Trusted Publisher (OIDC) for secure releases

### v0.3.0 — Stabilization (Current)

- Formal SQL subset specification ([`docs/SQL_SPEC.md`](SQL_SPEC.md)) with EBNF grammar
- Parser golden tests for all statement families (SELECT, INSERT, UPDATE, DELETE, DDL)
- Comprehensive test suite: **397 tests, 98% coverage**
- Parser fix: quoted strings with embedded spaces handled correctly
- Parser fix: escaped quotes (`''`, `""`) parsed correctly
- README restructured: limitations-first layout
- Microsoft Graph API engine: remote Excel files on OneDrive/SharePoint (experimental)

## Future

### Planned

- **DISTINCT**: Remove duplicate rows from SELECT results
- **OFFSET**: Pagination support (currently only LIMIT is supported)
- **Aggregate functions**: COUNT, SUM, AVG, MIN, MAX
- **GROUP BY**: Grouping with aggregate functions
- **Subqueries**: Nested SELECT statements
- **Multi-sheet JOIN**: Cross-sheet queries (INNER JOIN, LEFT JOIN)
- **Polars engine**: Optional backend using Polars instead of pandas
- **Async support**: Asyncio-compatible driver

### Not Planned

These are explicitly out of scope for excel-dbapi:

- Full ACID transactions (Excel files are not a database)
- Concurrent write support (single-writer model by design)
- ALTER TABLE / schema migration
- Stored procedures or triggers
- Foreign key enforcement

## Versioning

excel-dbapi follows [Semantic Versioning](https://semver.org/):

- **PATCH** (0.x.**y**): Bug fixes
- **MINOR** (0.**x**.0): New features, backward-compatible
- **MAJOR** (**x**.0.0): Breaking changes, stable API

**Current status**: Beta (0.x.x) — API may change before 1.0.0.

---

See [CHANGELOG.md](../CHANGELOG.md) for detailed release history.
