<!-- insertion marker -->
<a name="Unreleased"></a>
## Unreleased ([compare](https://github.com/yeongseon/excel-dbapi/compare/v0.1.3...HEAD))

### Added

- INSERT/UPDATE/DELETE support with `executemany()`.
- CREATE TABLE and DROP TABLE support.
- Pandas engine alongside openpyxl.
- WHERE (AND/OR, comparison operators), ORDER BY, and LIMIT.
- Parameter binding for `?` placeholders.
- Transaction simulation with autocommit, commit, and rollback.
- Cursor metadata: `description`, `rowcount`, `lastrowid`, `fetchmany()`.
- Expanded test coverage and usage examples.

### Changed

- SELECT results return tuples (DB-API style) instead of dicts.
- Minimum supported Python version is now 3.10.
- SQL errors now raise PEP 249 exceptions (ProgrammingError/NotSupportedError).
- `rollback()` is disabled when autocommit is enabled.

<a name="v2.0.0"></a>
## v2.0.0 (2025-04-03)

### Withdrawn

- Published in error. This release is withdrawn; continue using the 0.x series.

---

<a name="v0.1.3"></a>
## [v0.1.3](https://github.com/yeongseon/excel-dbapi/compare/v0.1.2...v0.1.3) (2025-04-03)

### Fixed

- Internal code cleanup and minor fixes

---

<a name="v0.1.2"></a>
## [v0.1.2](https://github.com/yeongseon/excel-dbapi/compare/v0.1.0...v0.1.2) (2025-04-03)

### Fixed

- Bug fixes and minor improvements

---

<a name="v0.1.0"></a>
## [v0.1.0](https://github.com/yeongseon/excel-dbapi/compare/96fbe280d7ce9e031d2df94ea950fed99ba1d283...v0.1.0) (2025-03-29)

### Initial Release

- Basic SELECT query support
- Cursor interface implementation
- PEP 249 compliant exception hierarchy
