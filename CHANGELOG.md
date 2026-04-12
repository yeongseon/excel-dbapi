# Changelog

All notable changes to this project will be documented in this file.

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
