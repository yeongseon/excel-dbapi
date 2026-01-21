# Product Requirements Document (PRD)
## excel-dbapi: Excel-as-Database DB-API Driver

**Version:** 1.0
**Date:** 2026-01-21
**Status:** Draft

---

## 1. Product Overview

excel-dbapi is a lightweight, Python DB-API 2.0 compliant driver that enables SQL-like querying of Excel files. The library treats Excel sheets as database tables, allowing data analysts and developers to use familiar SQL syntax for data extraction from spreadsheets without complex data transformation pipelines.

**Key Value Proposition:**
- Eliminate need for Excel-to-CSV conversion or database migration for simple querying
- Provide DB-API 2.0 standard interface for seamless integration with existing Python data workflows
- Support multiple execution engines for different use cases (performance vs. flexibility)

---

## 2. Goals

### 2.1 Primary Goals
- Provide a **PEP 249 (DB-API 2.0) compliant interface** for querying Excel files using SQL
- Enable **read-only SELECT operations** with support for filtering (WHERE clause)
- Support **multiple execution engines** (openpyxl for performance, pandas for flexibility)
- Maintain **backward compatibility** and stable API as the project matures

### 2.2 Secondary Goals
- Facilitate data exploration and analysis directly on Excel files
- Reduce development time for ad-hoc reporting and data extraction tasks
- Provide a stepping stone for migration from Excel-based workflows to proper database systems

---

## 3. Scope

### 3.1 Current Capabilities (v0.4.x)
- **Engine Support:** `openpyxl` and `pandas` engines
- **SQL Support:** SELECT, INSERT, UPDATE, DELETE
- **Query Features:** WHERE with AND/OR and comparison operators, ORDER BY, LIMIT
- **DDL Support:** CREATE TABLE, DROP TABLE
- **DB-API Interface:** Standard `connection` and `cursor` objects with:
  - `cursor()` method
  - `execute()` and `executemany()`
  - `fetchone()`, `fetchmany()`, `fetchall()`
  - `description`, `rowcount`, `lastrowid`
  - Context manager support (`with` statement)
- **Parameter Binding:** `?` placeholders for parameters
- **Transactions:** `commit()` and `rollback()` with optional autocommit
- **Exception Handling:** PEP 249 compliant exception hierarchy
- **Project Scaffolding:** Core tests and project structure (`src/`, `tests/`)

### 3.2 Planned Features (Future Releases)

#### Phase 1: Write Operations & DDL (v0.2.x) - Completed
- INSERT INTO statements (column-specified and column-unspecified)
- `executemany()` for bulk inserts
- CREATE TABLE (creates new worksheet with headers)
- DROP TABLE (removes worksheet)
- Auto-commit for write operations

#### Phase 2: Data Modification (v0.3.x) - Completed
- UPDATE statements with simple WHERE conditions
- DELETE statements (conditional and full sheet deletion)
- In-memory transaction simulation (commit/rollback)
- Track `rowcount` and `lastrowid`

#### Phase 3: Advanced SQL (v0.4.x) - Completed
- ORDER BY clause support
- LIMIT clause support
- Extended WHERE conditions (AND, OR, comparison operators)
- Improved SQL parser robustness

#### Phase 4: Production Readiness (v1.0.0)
- Complete read/write feature set
- Comprehensive documentation and examples
- Full test coverage (>90%)
- PyPI deployment and CI/CD pipeline

#### Phase 5: Extensibility (v2.0.x)
- SQLAlchemy Dialect integration
- Multi-sheet JOIN support
- Polars engine option
- Asynchronous query support
- Performance optimization for large files

### 3.3 Out of Scope (Non-Goals)
- Full-featured database capabilities (no stored procedures, triggers, views, indexes)
- Multi-file operations or federated queries across multiple Excel files
- Real-time collaboration or concurrent access control
- Excel-specific features (formulas, charts, formatting, pivot tables)
- Support for non-Excel file formats (CSV, ODS, etc.)
- Complex joins between different sheets (initial releases)
- Network protocol-based access (initial releases)

---

## 4. User Stories

### 4.1 Primary Users

**US-1: Data Analyst**  
*As a data analyst,*  
*I want to query Excel files using SQL,*  
*So that I can quickly extract and analyze data without exporting to CSV or importing into a database.*

**US-2: Python Developer**  
*As a Python developer,*  
*I want to use DB-API 2.0 interface for Excel files,*  
*So that I can integrate Excel data sources with existing database abstraction layers (SQLAlchemy).*

**US-3: QA Engineer**  
*As a QA engineer,*  
*I want to verify test data stored in Excel using SQL queries,*  
*So that I can automate validation checks without manual spreadsheet inspection.*

**US-4: Business User (Technical)**  
*As a technical business user,*  
*I want to generate ad-hoc reports from Excel files using SQL,*  
*So that I can quickly answer business questions without learning pandas or openpyxl APIs.*

### 4.2 Engine-Specific Stories

**US-5: Performance-Oriented User**  
*As a user processing large Excel files,*  
*I want to use the openpyxl engine for fast read operations,*  
*So that I can minimize memory usage and improve query performance.*

**US-6: Flexibility-Oriented User**  
*As a user needing data manipulation,*  
*I want to use the pandas engine for DataFrame compatibility,*  
*So that I can leverage pandas' rich ecosystem for further processing.*

---

## 5. Functional Requirements

### 5.1 Core Functionality

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-1 | Connection must accept Excel file path as parameter | P0 |
| FR-2 | Connection must support engine selection (openpyxl/pandas) | P0 |
| FR-3 | Connection must provide context manager interface | P0 |
| FR-4 | Cursor must execute SELECT queries | P0 |
| FR-5 | Cursor must support fetchone() and fetchall() | P0 |
| FR-6 | First row of each sheet must be treated as column headers | P0 |
| FR-7 | SQL parser must handle basic SELECT statements with WHERE clause | P0 |
| FR-8 | Exception hierarchy must follow PEP 249 standard | P0 |
| FR-9 | pandas engine must support save() method | P0 |
| FR-10 | Table name resolution must map to sheet names (case-sensitive) | P0 |

### 5.2 Future Requirements (Planned)

| ID | Requirement | Phase |
|----|-------------|-------|
| FR-11 | Support INSERT INTO statements | Phase 1 |
| FR-12 | Support executemany() for bulk operations | Phase 1 |
| FR-13 | Support CREATE TABLE (new worksheet creation) | Phase 1 |
| FR-14 | Support DROP TABLE (worksheet removal) | Phase 1 |
| FR-15 | Support UPDATE statements | Phase 2 |
| FR-16 | Support DELETE statements | Phase 2 |
| FR-17 | Implement transaction simulation (commit/rollback) | Phase 2 |
| FR-18 | Track rowcount and lastrowid attributes | Phase 2 |
| FR-19 | Support ORDER BY clause | Phase 3 |
| FR-20 | Support LIMIT clause | Phase 3 |
| FR-21 | Support extended WHERE conditions (AND, OR, operators) | Phase 3 |
| FR-22 | Support parameter binding (?) placeholders | Phase 3 |
| FR-23 | Provide SQLAlchemy Dialect | Phase 5 |
| FR-24 | Support multi-sheet JOIN operations | Phase 5 |

### 5.3 Error Handling

| ID | Requirement |
|----|-------------|
| FR-25 | Raise `DatabaseError` for SQL syntax errors |
| FR-26 | Raise `InterfaceError` for invalid connection usage |
| FR-27 | Raise `OperationalError` for file access issues (missing file, permissions) |
| FR-28 | Raise `DataError` for data type mismatches |
| FR-29 | Raise `ProgrammingError` for invalid table/column names |
| FR-30 | Raise `NotSupportedError` for unsupported SQL features |

---

## 6. Non-Functional Requirements

### 6.1 Performance
| ID | Requirement | Target |
|----|-------------|--------|
| NFR-1 | openpyxl engine must support files up to 10MB with <2s initial load | <2s |
| NFR-2 | pandas engine must support files up to 50MB | Memory <500MB |
| NFR-3 | SELECT query execution must complete within 1s for <10,000 rows | <1s |
| NFR-4 | Memory footprint must be optimized for read-only operations | <2x file size |

### 6.2 Compatibility
| ID | Requirement |
|----|-------------|
| NFR-5 | Must support Python 3.10+ |
| NFR-6 | Must support openpyxl 3.1.0+ |
| NFR-7 | Must support pandas 2.0.0+ |
| NFR-8 | Must pass PEP 249 compliance tests |
| NFR-9 | Must support .xlsx file format (Excel 2007+) |

### 6.3 Quality
| ID | Requirement | Target |
|----|-------------|--------|
| NFR-10 | Code coverage must exceed 80% for core modules | >80% |
| NFR-11 | All public APIs must have type hints | 100% |
| NFR-12 | Must pass static type checking (mypy) | Zero errors |
| NFR-13 | Must pass linting (ruff, black, isort) | Zero warnings |
| NFR-14 | Security scanning (bandit) must pass | Zero issues |

### 6.4 Reliability
| ID | Requirement |
|----|-------------|
| NFR-15 | Must handle malformed Excel files gracefully |
| NFR-16 | Must handle empty sheets without crashing |
| NFR-17 | Must preserve original file on write failures |
| NFR-18 | Must provide clear error messages for common issues |

### 6.5 Usability
| ID | Requirement |
|----|-------------|
| NFR-19 | API must be discoverable (follows DB-API 2.0 conventions) |
| NFR-20 | Error messages must be actionable |
| NFR-21 | Documentation must include quick start guide |
| NFR-22 | Examples must cover common use cases |

---

## 7. Milestones

### 7.1 Current Release
**v0.4.x** (Current)
- ✅ Engines: openpyxl and pandas
- ✅ SELECT, INSERT, UPDATE, DELETE
- ✅ WHERE (AND/OR, comparison operators), ORDER BY, LIMIT
- ✅ CREATE TABLE, DROP TABLE
- ✅ `executemany()`, `fetchmany()`, `description`, `rowcount`, `lastrowid`
- ✅ Parameter binding (`?` placeholders)
- ✅ Transaction simulation (commit/rollback)
- ✅ Tests and documentation updates

### 7.2 Upcoming Releases

**Milestone 1: Write Operations & DDL (v0.2.x)**
- **Status:** Completed

**Milestone 2: Data Modification (v0.3.x)**
- **Status:** Completed

**Milestone 3: Advanced SQL Features (v0.4.x)**
- **Status:** Completed

**Milestone 4: Production Release (v1.0.0)**
- **Deliverables:**
  - Feature-complete read/write support
  - Comprehensive documentation (usage, API, examples)
  - >90% test coverage
  - PyPI deployment
  - CI/CD pipeline (GitHub Actions)
  - Changelog and release notes
- **Success Criteria:** Stable API, zero breaking changes from beta to 1.0

**Milestone 5: Advanced Features (v2.0.x)**
- **Deliverables:**
  - SQLAlchemy Dialect
  - Multi-sheet JOIN support
  - Polars engine option
  - Asynchronous query support (optional)
  - Performance optimizations for large files
- **Success Criteria:** SQLAlchemy integration passes standard tests

---

## 8. Risks and Mitigation

### 8.1 Technical Risks

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| **Excel file format limitations** | Medium | High | Document file size and complexity limits; provide clear error messages |
| **SQL parser complexity** | High | High | Start with simple subset; use existing parsing libraries where possible |
| **Performance degradation on large files** | High | Medium | Implement streaming reads; provide performance benchmarks; recommend pandas for complex operations |
| **Write operation data corruption** | Medium | High | Implement atomic writes; create backup before modifications; extensive testing |
| **Transaction simulation inconsistency** | Medium | High | Clearly document limitations; follow DB-API 2.0 transaction semantics |

### 8.2 Project Risks

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| **Scope creep** | High | Medium | Maintain strict roadmap; defer advanced features to v2.0+ |
| **Maintainer bandwidth** | Medium | High | Encourage community contributions; keep codebase modular |
| **Breaking changes** | Medium | Medium | Semantic versioning; deprecation warnings; clear migration guides |
| **Low adoption** | Low | Medium | Improve documentation; provide compelling examples; integrate with popular tools |

### 8.3 Dependency Risks

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| **openpyxl/pandas API changes** | Low | High | Pin dependency versions; monitor upstream releases; automated testing |
| **Python version incompatibility** | Low | Medium | Support multiple Python versions; test on all supported versions |
| **PyPI package naming conflict** | Low | Medium | Verify package name availability early; establish unique branding |

---

## 9. Versioning Policy Recommendation

### 9.1 Semantic Versioning (SemVer)

excel-dbapi should follow **Semantic Versioning 2.0.0**:
- **MAJOR** version: Incompatible API changes
- **MINOR** version: Backwards-compatible functionality additions
- **PATCH** version: Backwards-compatible bug fixes

**Format:** `MAJOR.MINOR.PATCH` (e.g., `1.0.0`, `2.1.3`)

### 9.2 Version Lifecycle

#### Alpha Phase (0.1.x)
- Experimental features
- Rapid iteration
- **No stability guarantees**
- Breaking changes expected

#### Beta Phase (0.2.x - 0.4.x)
- Feature-complete for phase goals
- API stabilization
- Breaking changes may occur with deprecation warnings
- Community feedback encouraged

#### Release Candidate (1.0.0-rc.x)
- Feature complete for v1.0
- Bug fixes only
- No breaking changes
- Final testing phase

#### Stable Release (1.x.x)
- Backwards-compatible API
- Documented deprecation process
- Long-term support (LTS) for major versions

### 9.3 Release Criteria

#### Patch Release (x.x.Z)
- Bug fixes only
- No API changes
- Changelog entry required

#### Minor Release (x.Y.0)
- New backwards-compatible features
- Deprecation of old features (with warnings)
- Changelog and migration guide if deprecations

#### Major Release (X.0.0)
- Breaking API changes
- Comprehensive migration guide
- Documentation updates
- Support period for previous major version (6 months)

### 9.4 Deprecation Policy

1. **Announce** deprecation in release notes
2. **Add warnings** when deprecated feature is used
3. **Maintain** for at least **one minor version**
4. **Remove** in next major version

**Example:**
- v1.2.0: Feature X deprecated, warnings added
- v1.3.0: Feature X still available with warnings
- v2.0.0: Feature X removed

### 9.5 Support Policy

- **Latest major version:** Active development and security patches
- **Previous major version:** Security patches only (6 months after new major release)
- **Older versions:** No support

---

## 10. Appendices

### A. Glossary

| Term | Definition |
|------|------------|
| DB-API 2.0 | Python database API specification (PEP 249) |
| Sheet | Excel worksheet equivalent to a database table |
| Engine | Underlying implementation (openpyxl or pandas) that executes queries |
| PEP | Python Enhancement Proposal |
| Semantic Versioning | Version numbering scheme (MAJOR.MINOR.PATCH) |

### B. References

- [PEP 249 - Python Database API Specification v2.0](https://www.python.org/dev/peps/pep-0249/)
- [openpyxl Documentation](https://openpyxl.readthedocs.io/)
- [pandas Documentation](https://pandas.pydata.org/)
- [Semantic Versioning 2.0.0](https://semver.org/)

### C. Architecture Overview (High-Level)

```
┌─────────────────────────────────────────┐
│         User Application                 │
│         (SQL Queries)                   │
└──────────────┬──────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────┐
│      excel-dbapi (DB-API 2.0 Layer)    │
│  ┌──────────┐      ┌──────────────┐    │
│  │Connection│─────▶│   Cursor     │    │
│  └──────────┘      └──────┬───────┘    │
│                            │            │
│                            ▼            │
│                    ┌───────────────┐   │
│                    │ SQL Parser    │   │
│                    └───────┬───────┘   │
└────────────────────────────┼───────────┘
                             │
                ┌────────────┴────────────┐
                ▼                         ▼
        ┌───────────────┐       ┌───────────────┐
        │  OpenpyxlEngine │     │  PandasEngine  │
        │  (Read-only)   │     │  (Read/Write)  │
        └───────┬───────┘       └───────┬───────┘
                │                       │
                └───────────┬───────────┘
                            ▼
                    ┌───────────────┐
                    │  Excel File   │
                    │  (.xlsx)      │
                    └───────────────┘
```

---

**Document History**

| Version | Date | Author | Changes |
|---------|------|--------|---------|
| 1.0 | 2026-01-21 | PRD Author | Initial PRD creation aligned to v0.1.x roadmap |
| 1.1 | 2026-01-21 | PRD Author | Updated to reflect v0.4.x completion |
