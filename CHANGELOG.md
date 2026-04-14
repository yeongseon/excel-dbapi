# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [0.4.1] - 2026-04-13

### Added
- RIGHT JOIN and RIGHT OUTER JOIN support
- Chained (multi-table) JOIN support with iterative fold execution
- FULL OUTER JOIN and CROSS JOIN support
- SELECT * with JOIN: wildcard expansion to all qualified columns
- GROUP BY and aggregates in JOIN queries
- Multi-row INSERT: `INSERT INTO t VALUES (1, 'a'), (2, 'b')`
- INSERT...SELECT: `INSERT INTO t2 SELECT col1, col2 FROM t1 WHERE ...`
- Compound SELECT set operations: `UNION`, `UNION ALL`, `INTERSECT`, `EXCEPT`
- Column aliases: `SELECT col AS alias`
- COUNT(DISTINCT col) support
- Multi-column ORDER BY
- NOT operator, parenthesized WHERE, NOT IN/LIKE/BETWEEN
- Arithmetic expressions in SELECT columns
- CASE WHEN expressions in SELECT, UPDATE, and WHERE
- UPSERT: `INSERT ... ON CONFLICT DO NOTHING/UPDATE`
- ALTER TABLE: ADD COLUMN, DROP COLUMN, RENAME COLUMN
- Window functions: ROW_NUMBER, RANK, DENSE_RANK with PARTITION BY and ORDER BY
- Aggregate FILTER clause
- ILIKE operator and ESCAPE clause for pattern matching
- Common Table Expressions (CTE): `WITH name AS (SELECT ...) SELECT ...`
- EXISTS, NOT EXISTS, and scalar subqueries
- Subqueries in UPDATE/DELETE WHERE clauses
- Scalar functions: ABS, ROUND, REPLACE
- PEP 249 type constructors (Date, Time, Timestamp, Binary, STRING, BINARY, NUMBER, DATETIME, ROWID)
- File-based locking for concurrent write protection
- Configurable row limits and large-workbook safeguards
- SharePoint/OneDrive DSN locators (`sharepoint://`, `onedrive://`)
- Graph engine: configurable timeout, retry backoff, auth diagnostics
- Graph engine: eTag-based concurrent write conflict detection
- Graph engine: optimized UPDATE/DELETE with targeted range patches
- Packaging smoke tests in CI (wheel and sdist install verification)
- Property-based and fuzz tests for parser and tokenizer
- Test count: 546 → 1264

### Changed
- SQL three-valued logic (TRUE/FALSE/UNKNOWN) for WHERE, HAVING, and ON evaluation
- NOT IN with NULL follows SQL standard semantics
- Casefold-based column and sheet name resolution (case-insensitive matching)
- Pandas engine validates duplicate headers and enforces data_only mode
- Reflection type inference is now configurable and more accurate
- Date, datetime, and boolean values sort by native type
- JOIN source collision checks track all previously introduced table references
- SQL grammar and specification updated for all new syntax

### Fixed
- DB-API exception wrapping at connect, execute, commit, rollback, and close time
- BadZipFile and backend construction errors wrapped as OperationalError
- Cursor state reset in executemany; executemany always snapshots before batch
- Negative LIMIT/OFFSET values now rejected with clear error
- LIKE with NULL operand correctly returns UNKNOWN
- Window expressions collected from CASE and WHERE nodes
- Parameter validation rejects execute with params but no placeholders
- Guard fetch operations on missing result set
- Metadata table protected from DDL operations
- Graph worksheet IDs URL-encoded in API calls
- Header whitespace trimming with duplicate detection
- Compound query placeholder counting for comma-attached tokens

## [0.4.0] - 2026-04-13

### Added
- Multi-sheet JOIN support: INNER JOIN and LEFT JOIN with alias syntax
- JOIN parser: qualified column references (`a.id`), ON condition parsing, alias support
- JOIN executor: hash join algorithm with namespaced row flattening
- LEFT JOIN: NULL fill for unmatched right-table rows
- WHERE, ORDER BY, LIMIT, OFFSET work with JOIN queries
- 34 new JOIN tests (20 parser, 9 executor, 5 boundary)
- SQL_SPEC.md updated with JOIN syntax, grammar, and constraints

### Changed
- Parser now always emits `from` and `joins` keys in SELECT AST
- Test count: 397 → 546
- Version bumped to 0.4.0

## [0.3.0] - 2026-04-12

### Added
- Formal SQL subset specification (`docs/SQL_SPEC.md`) with EBNF grammar
- Parser golden tests for all statement families (SELECT, INSERT, UPDATE, DELETE, DDL)
- Reflection helpers unit tests
- Comprehensive low-coverage module tests (executor, backends, graph engine)

### Changed
- README restructured: limitations-first layout, Graph API moved to experimental section
- Test coverage: 84% → 98% (397 tests)

### Fixed
- Parser tokenizer: quoted strings with embedded spaces now handled correctly
- Parser: escaped quotes (`''`, `""`) parsed correctly in all contexts

## [0.2.1] - 2026-04-12

### Added
- Project logo (modern minimalist SVG)
- Comprehensive README documentation (WHERE operators section, Related Projects)
- Contributing guide, Code of Conduct, Security and Support policies
- Development tooling: Makefile, .editorconfig, pre-commit-config, codecov.yml, git-cliff config
- GitHub issue/PR templates and project management files
- py.typed marker for PEP 561 compliance
- twine check step in publish workflow

### Changed
- Classifier updated from Alpha to Beta
- Homepage metadata updated
- Version bumped to 0.2.1

### Fixed
- Rollback documentation, absolute logo URLs, metadata alignment

## [0.2.0] - 2026-04-12

### Added
- IN, BETWEEN, LIKE operators for SQL parser and executor
- Test coverage reporting with Codecov CI integration

### Changed
- Version bumped from 0.1.x to 0.2.0 (skipping reserved PyPI versions)

### Fixed
- All mypy strict errors resolved; strict mode enabled in CI
- CI: install pandas extra for test suite

## [0.1.0] - 2026-04-12

### Added
- PEP 249 (DB-API 2.0) compliant driver for Excel files
- SQL support: SELECT, INSERT, UPDATE, DELETE, CREATE TABLE, DROP TABLE
- WHERE clauses with AND/OR, comparison operators, LIKE, IN, IS NULL/IS NOT NULL
- ORDER BY and LIMIT for SELECT queries
- Openpyxl engine (default) for local .xlsx files
- Pandas engine (optional) for DataFrame-based operations
- Microsoft Graph API engine (optional) for remote Excel files
- Formula injection defense (enabled by default)
- Transaction simulation (commit/rollback)
- Reflection helpers for dialect integration
- Metadata sheet support for schema persistence
