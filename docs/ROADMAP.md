# excel-dbapi Project Roadmap

This document outlines the planned milestones and tasks for the `excel-dbapi` project.  
The project aims to implement a **PEP 249 (DB-API 2.0) compliant driver for Excel files.**

---

## Version 0.1.x - Basic Read-Only Support

**Goal:** Provide minimal functionality to read Excel data using DB-API 2.0 standard.

### Completed
- [x] Implement `ExcelConnection` class
  - [x] `cursor()` method
  - [x] `close()` method
  - [x] Context manager support (`__enter__`, `__exit__`)
  - [x] Engine selection (currently openpyxl only)
- [x] Implement `ExcelCursor` class
  - [x] `execute()` method (SELECT query parser & executor)
  - [x] `fetchone()` and `fetchall()` methods
  - [x] `close()` method
- [x] Implement basic SQL parser (SELECT structure only)
- [x] Sheet-to-Table mapping (first row as header)
- [x] Implement PEP 249 Exception hierarchy
- [ ] Parameter binding support (`?` placeholder) (⚠️ Partial: parsing not applied yet)
- [x] Basic unit tests (Connection, Cursor, Parser, Executor)
- [x] Add usage examples (README.md + examples/basic_usage.py)
- [x] Project structure reorganization (`src/`, `tests/`)
- [x] pyproject.toml modernization

---

## Version 0.2.x - Write Operations & Basic DDL

**Goal:** Add data write capability and minimal DDL operations.

### Completed
- [x] Implement `INSERT INTO` execution
  - [x] Column-specified and column-unspecified INSERT
  - [x] `executemany()` support
- [x] Add row(s) to Excel sheet
- [x] Save changes to file (auto-commit)
- [x] Implement basic DDL
  - [x] `CREATE TABLE` → Create new worksheet with headers
  - [x] `DROP TABLE` → Remove worksheet (optional)
- [x] Unit tests for INSERT and DDL
- [x] Documentation update

---

## Version 0.3.x - UPDATE, DELETE, Transaction Simulation

**Goal:** Enable data modification and transaction simulation.

### Completed
- [x] Implement `UPDATE` execution
  - [x] Simple `WHERE` condition support
- [x] Implement `DELETE` execution
  - [x] Conditional and full sheet deletion
- [x] In-memory transaction simulation
  - [x] `commit()` and `rollback()` behavior
- [x] Track `rowcount` and `lastrowid`
- [x] Unit tests for UPDATE, DELETE, transactions
- [x] Documentation update

---

## Version 0.4.x - SQL Features Expansion

**Goal:** Enhance SQL support for advanced queries.

### Completed
- [x] Implement `ORDER BY`, `LIMIT` clauses
- [x] Extend `WHERE` condition (AND, OR, comparison operators)
- [x] Improve parser robustness
- [x] Comprehensive unit tests
- [x] Advanced usage examples in documentation

---

## Version 1.0.0 - Production Ready Release

**Goal:** Stabilize and release v1.0.0

### Completed
- [x] Complete feature set for read/write
- [x] Extensive documentation & examples
- [x] Full test coverage
- [x] Packaging, versioning, CI/CD ready

### Pending
- [ ] PyPI deployment

---

## Version 2.0.x - Advanced & Extensibility

**Optional future plan**

- [ ] SQLAlchemy Dialect Integration
- [ ] Multi-sheet JOIN support
- [ ] Polars Engine support (optional)
- [ ] Asynchronous query support
- [ ] Performance optimization for large files

---
