# Changelog

All notable changes to this project will be documented in this file.

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
- Oracle review findings: rollback docs, absolute logo URLs, metadata alignment

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
