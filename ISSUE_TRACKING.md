# Issue Tracking Checklist

Quick reference checklist for tracking progress on identified code review issues.

## Phase 1: Critical Security Fixes (IMMEDIATE)

### Formula Injection Vulnerabilities
- [ ] **Issue #1**: Fix formula injection in OpenpyxlExecutor (`src/excel_dbapi/engine/openpyxl_executor.py:178`)
  - Add `sanitize_cell_value()` function
  - Apply to INSERT operations
  - Apply to UPDATE operations
  - Add test cases

- [ ] **Issue #2**: Fix formula injection in PandasExecutor (`src/excel_dbapi/engine/pandas_executor.py:129`)
  - Add `sanitize_cell_value()` function
  - Apply to INSERT operations
  - Apply to UPDATE operations
  - Add test cases

## Phase 2: PEP 249 Compliance (HIGH PRIORITY)

### Required Methods and Exports
- [ ] **Issue #6**: Add missing cursor methods (`src/excel_dbapi/cursor.py`)
  - Add `setinputsizes(sizes)` method
  - Add `setoutputsize(size, column=None)` method
  - Add test cases

- [ ] **Issue #7**: Export exception classes (`src/excel_dbapi/__init__.py`)
  - Import all exception classes
  - Add to `__all__` list
  - Add test for module-level access

- [ ] **Issue #8**: Fix exception types (`src/excel_dbapi/connection.py:38`)
  - Change InterfaceError to OperationalError for engine errors
  - Review all exception usage
  - Update tests

- [ ] **Issue #9**: Add file existence validation (`src/excel_dbapi/connection.py:25`)
  - Check file exists before loading
  - Raise OperationalError if not found
  - Add test cases

## Phase 3: Security Hardening (HIGH PRIORITY)

### Path and File Security
- [ ] **Issue #3**: Add path validation (`src/excel_dbapi/connection.py:29`)
  - Use `os.path.realpath()` to resolve paths
  - Optional: Add directory restrictions
  - Document security considerations
  - Add test cases

- [ ] **Issue #4**: Secure temp file in OpenpyxlEngine (`src/excel_dbapi/engine/openpyxl_engine.py:56`)
  - Add `os.chmod(temp_file, 0o600)`
  - Document temp file behavior
  - Add test to verify permissions

- [ ] **Issue #5**: Secure temp file in PandasEngine (`src/excel_dbapi/engine/pandas_engine.py:30`)
  - Add `os.chmod(temp_file, 0o600)`
  - Document temp file behavior
  - Add test to verify permissions

## Phase 4: Bug Fixes (MEDIUM PRIORITY)

### Parsing and Logic Bugs
- [ ] **Issue #10**: Fix escaped quotes in strings (`src/excel_dbapi/engine/parser.py:29`)
  - Handle `\'` and `\"` escape sequences
  - Add test with escaped quotes
  - Test edge cases

- [ ] **Issue #11**: Fix/document AND/OR precedence (`src/excel_dbapi/engine/openpyxl_executor.py:222`)
  - Option A: Document limitation in README
  - Option B: Implement proper precedence
  - Add tests for mixed operators
  - Add NotSupportedError for unsupported cases

- [ ] **Issue #12**: Fix parameter handling in BaseEngine (`src/excel_dbapi/engine/base.py:39`)
  - Raise NotImplementedError if params provided
  - Update docstring
  - Add test case

- [ ] **Issue #13**: Improve table name error messages (`src/excel_dbapi/engine/executor.py:24`)
  - Preserve original table name
  - Use original in error messages
  - Add test case

- [ ] **Issue #14**: Fix None comparison handling (`src/excel_dbapi/engine/openpyxl_executor.py:263`)
  - Handle None explicitly before string conversion
  - Add test cases for NULL comparisons
  - Document SQL NULL semantics

## Phase 5: Code Quality (LOW PRIORITY)

### Cleanup and Improvements
- [ ] **Issue #15**: Remove unused import (`src/excel_dbapi/engine/pandas_engine.py:1`)
  - Remove `from copy import deepcopy`

- [ ] **Issue #16**: Add type annotations to decorators
  - `src/excel_dbapi/connection.py:12`
  - `src/excel_dbapi/cursor.py:9`
  - Use `@wraps` from functools
  - Add proper typing

- [ ] **Issue #17**: Add duplicate column validation (`src/excel_dbapi/engine/parser.py:236`)
  - Check for duplicate column names
  - Raise ValueError if found
  - Add test case

- [ ] **Issue #18**: Add LIMIT validation (`src/excel_dbapi/engine/parser.py:148`)
  - Set MAX_LIMIT = 1048576
  - Validate LIMIT <= MAX_LIMIT
  - Add test cases

- [ ] **Issue #19**: Add empty params check (`src/excel_dbapi/cursor.py:53`)
  - Handle empty `seq_of_params` explicitly
  - Add test case

## Phase 6: Enhancements (LOW PRIORITY)

### Feature Additions
- [ ] **Issue #20**: Add cursor context manager (`src/excel_dbapi/cursor.py`)
  - Add `__enter__` method
  - Add `__exit__` method
  - Add test using `with` statement

- [ ] **Issue #21**: Add comprehensive docstrings
  - ExcelConnection methods
  - ExcelCursor methods
  - Engine classes
  - Follow PEP 257

- [ ] **Issue #22**: Document memory limitations (`src/excel_dbapi/cursor.py:28`)
  - Add note about result storage
  - Document Excel row limits
  - Update README

- [ ] **Issue #23**: No action needed (Warning class is correct per PEP 249)

---

## Progress Summary

**Completed:** 0 / 23  
**In Progress:** 0 / 23  
**Blocked:** 0 / 23  

### By Phase
- Phase 1 (Critical Security): 0 / 2
- Phase 2 (PEP 249 Compliance): 0 / 4
- Phase 3 (Security Hardening): 0 / 3
- Phase 4 (Bug Fixes): 0 / 5
- Phase 5 (Code Quality): 0 / 5
- Phase 6 (Enhancements): 0 / 3

---

## Notes

- Each issue should be fixed in a separate, small commit
- Add test cases for each fix
- Run full test suite after each fix
- Update CHANGELOG.md for user-facing changes
- Consider security implications of each change

---

*Last Updated: 2026-01-24*  
*Use this checklist to track progress as issues are resolved*
