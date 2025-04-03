# excel-dbapi Project Roadmap

This document outlines the planned milestones and tasks for the `excel-dbapi` project.  
The project aims to implement a DBAPI-compliant connector for Excel files.

---

## Version 0.1.x - Basic Read-Only Support

**Goal:** Provide minimum functionality to read Excel data using DBAPI 2.0 standard.

### TODO
- [x] Implement `connect()` function and `Connection` class
  - [x] `cursor()` method
  - [x] `close()` method
- [ ] Implement `Cursor` class
  - [ ] `execute()` method (SELECT query parser)
  - [ ] `fetchone()` and `fetchall()` methods
  - [ ] `description` attribute
  - [ ] `rowcount` and `lastrowid` default values
  - [ ] `close()` method
- [ ] Sheet-to-Table mapping (first row as header)
- [ ] Basic exception handling (OperationalError, ProgrammingError, etc.)
- [ ] Unit tests for SELECT functionality
- [ ] Add usage examples in `README.md`

---

## Version 0.2.x - Write Operations & Basic DDL

**Goal:** Support data insertion and minimal DDL operations.

### TODO
- [ ] Implement `INSERT INTO` query parser
  - [ ] Column-specified and column-unspecified INSERT
  - [ ] Support for `executemany()`
- [ ] Add row(s) to Excel sheet
- [ ] Save changes to the file (Auto-commit mode)
- [ ] Implement basic DDL
  - [ ] `CREATE TABLE` (create new worksheet with headers)
  - [ ] `DROP TABLE` (remove worksheet) (optional)
- [ ] Unit tests for INSERT and DDL
- [ ] Update documentation

---

## Version 0.3.x - UPDATE, DELETE, Transaction Support

**Goal:** Enable data modification and transactional control.

### TODO
- [ ] Implement `UPDATE` query parser and execution
  - [ ] Simple `WHERE` condition support
- [ ] Implement `DELETE` query parser and execution
  - [ ] Conditional and full table deletion
- [ ] Implement transaction control
  - [ ] `commit()` and `rollback()` methods
  - [ ] In-memory buffering
- [ ] Track `rowcount` for UPDATE & DELETE
- [ ] Unit tests for UPDATE, DELETE, transaction
- [ ] Update documentation

---

## Version 0.4.x - Advanced SQL & DDL Completion

**Goal:** Expand SQL capabilities and complete DDL support.

### TODO
- [ ] Implement `JOIN` queries (INNER JOIN)
- [ ] Implement aggregation functions
  - [ ] `COUNT()`, `SUM()`, `AVG()`, `MIN()`, `MAX()`
  - [ ] `GROUP BY` support
- [ ] Implement `ORDER BY` clause
- [ ] Extend `WHERE` condition parser (AND, OR, comparisons)
- [ ] Finalize `CREATE TABLE` and `DROP TABLE`
- [ ] Improve SQL parser
- [ ] Add comprehensive tests
- [ ] Update documentation with advanced examples

