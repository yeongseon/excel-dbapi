<!-- insertion marker -->
<a name="Unreleased"></a>
## Unreleased ([compare](https://github.com/yeongseon/excel-dbapi/compare/v2.0.0...HEAD))

_No unreleased changes yet._

<!-- insertion marker -->
<a name="v2.0.0"></a>
## [v2.0.0](https://github.com/yeongseon/excel-dbapi/compare/v0.1.3...v2.0.0) (2025-04-03)

### ✨ Added

- Engine selection support: `pandas` and `openpyxl`
- `ExcelConnection(engine=...)` option to choose engine
- `PandasEngine`, `OpenpyxlEngine` implementation
- Sheet-to-table mapping (using the first row as headers)
- `save()` method for `PandasEngine`
- Unit tests for engine structure and save feature
- Refactored project structure (introduced `engine` module)

### 🛠️ Changed

- Refactored `ExcelConnection` class to support engine injection
- SQL parser table name normalization (case-sensitive)
- Updated test cases for the new engine structure

### Removed

- Legacy modules: `api.py`, `query.py`, `table.py`, `remote.py`
- Removed `src/excel_dbapi/engines/` directory (merged into `engine` package)

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
