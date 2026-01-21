# Development Tasks (Target: v1.0.0)

This document tracks the agreed development scope through v1.0.0.
It follows the ROADMAP milestones and includes additional DB-API features:
- `Cursor.description`
- `Cursor.rowcount`, `Cursor.lastrowid`
- `Cursor.fetchmany()`
- Parameter binding (`?`) execution

## Phase 0: Foundations (pre-v0.2)
- Define execution result model (rows, description, rowcount, lastrowid).
- Align cursor/connection responsibilities for engine delegation.
- Ensure engine API supports parameterized execution.
- Add DB-API module constants in `__init__.py` (apilevel, threadsafety, paramstyle).
- Update tests for new cursor attributes and DB-API surface.

## v0.2.x: Write Operations & Basic DDL
- Implement INSERT (column-specified and column-omitted) for openpyxl.
- Implement INSERT for pandas.
- Implement `executemany()` for INSERT.
- Implement CREATE TABLE and DROP TABLE for openpyxl.
- Implement CREATE TABLE and DROP TABLE for pandas.
- Define and document auto-commit behavior.
- Add unit tests for INSERT/DDL and `executemany()`.
- Update docs/README for new write capabilities.

## v0.3.x: UPDATE, DELETE, Transaction Simulation
- Implement UPDATE with simple WHERE (openpyxl).
- Implement UPDATE with simple WHERE (pandas).
- Implement DELETE with simple WHERE and full table delete (openpyxl).
- Implement DELETE with simple WHERE and full table delete (pandas).
- Add transaction simulation (commit/rollback) at connection level.
- Track `rowcount` and `lastrowid` for write operations.
- Add unit tests for UPDATE/DELETE and transaction behavior.
- Update docs/README with transaction semantics.

## v0.4.x: SQL Features Expansion
- Implement ORDER BY and LIMIT (openpyxl).
- Implement ORDER BY and LIMIT (pandas).
- Extend WHERE to support AND/OR and comparison operators.
- Implement parameter binding execution for `?` placeholders.
- Harden parser error handling and messaging.
- Add tests for extended SQL parsing and execution.
- Add advanced usage examples in docs.

## v1.0.0: Production Readiness
- Confirm full DB-API cursor surface (`description`, `rowcount`, `lastrowid`, `fetchmany`).
- Verify read/write feature completeness across engines.
- Ensure documentation coverage (usage, API, examples) matches implementation.
- Achieve coverage targets and fix gaps.
- Stabilize public API and finalize release notes.

## Validation Checklist (apply to every phase)
- Tests: `make setup` (once), `make test` (per change).
- Docs: update `CHANGELOG.md`, `docs/PRD.md`, `docs/ROADMAP.md` as needed.
- Version: update `pyproject.toml` when releasing.
