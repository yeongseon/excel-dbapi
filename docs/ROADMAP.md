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
  - [x] Engine selection (`pandas`, `openpyxl`)
- [x] Implement `ExcelCursor` class
  - [x] `execute()` method (SELECT query parser & executor)
  - [x] `fetchone()` and `fetchall()` methods
  - [x] `close()` method
- [x] Implement basic SQL parser (SELECT & INSERT structure)
- [x] Sheet-to-Table mapping (first row as header)
- [x] Implement PEP 249 Exception hierarchy
- [x] Parameter binding support (`?` placeholder)
- [x] Basic unit tests (Connection, Cursor, Parser, Executor)
- [x] Add usage examples in `README.md`
- [x] Project structure reorganization (`src/`, `tests/`)
- [x] pyproject.toml modernization

---

## Version 0.2.x - Write Operations & Basic DDL

**Goal:** Add data write capability and minimal DDL operations.

### Planned
- [ ] Implement `INSERT INTO` execution
  - [ ] Column-specified and column-unspecified INSERT
  - [ ] `executemany()` support
- [ ] Add row(s) to Excel sheet
- [ ] Save changes to file (auto-commit)
- [ ] Implement basic DDL
  - [ ] `CREATE TABLE` → Create new worksheet with headers
  - [ ] `DROP TABLE` → Remove worksheet (optional)
- [ ] Unit tests for INSERT and DDL
- [ ] Documentation update

---

## Version 0.3.x - UPDATE, DELETE, Transaction Simulation

**Goal:** Enable data modification and transaction simulation.

### Planned
- [ ] Implement `UPDATE` execution
  - [ ] Simple `WHERE` condition support
- [ ] Implement `DELETE` execution
  - [ ] Conditional and full sheet deletion
- [ ] In-memory transaction simulation
  - [ ] `commit()` and `rollback()` behavior
- [ ] Track `rowcount` and `lastrowid`
- [ ] Unit tests for UPDATE, DELETE, transactions
- [ ] Documentation update

---

## Version 0.4.x - SQL Features Expansion

**Goal:** Enhance SQL support for advanced queries.

### Planned
- [ ] Implement `ORDER BY`, `LIMIT` clauses
- [ ] Extend `WHERE` condition (AND, OR, comparison operators)
- [ ] Improve parser robustness
- [ ] Comprehensive unit tests
- [ ] Advanced usage examples in documentation

---

## Version 1.0.0 - Production Ready Release

**Goal:** Stabilize and release v1.0.0

### Planned
- [ ] Complete feature set for read/write
- [ ] Extensive documentation & examples
- [ ] Full test coverage
- [ ] Packaging, versioning, CI/CD ready
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
