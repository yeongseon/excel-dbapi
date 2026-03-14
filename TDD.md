# TDD.md — excel-dbapi Technical Design Document

> **Version**: 1.0.0  
> **Last Updated**: 2026-03-15  
> **Status**: v1.0.0 Stable (feature-complete)

---

## 1. Implementation Overview

excel-dbapi v1.0.0 is fully implemented, tested, and published to PyPI. This document describes the detailed technical design of each module, the SQL grammar specification, engine abstraction patterns, and comprehensive usage examples.

### 1.1 Development Phases (All Complete)

| Phase | Version | Features | Status |
|-------|---------|----------|--------|
| Phase 1 | v0.1.x | SELECT, DB-API interface, openpyxl engine | ✅ |
| Phase 2 | v0.2.x | INSERT, CREATE TABLE, DROP TABLE, executemany | ✅ |
| Phase 3 | v0.3.x | UPDATE, DELETE, transaction simulation | ✅ |
| Phase 4 | v0.4.x | ORDER BY, LIMIT, extended WHERE (AND/OR) | ✅ |
| Phase 5 | v1.0.0 | Production release, documentation, tests | ✅ |

### 1.2 File Map

```
src/excel_dbapi/
├── __init__.py              30 lines   Module constants + connect()
├── connection.py            93 lines   ExcelConnection
├── cursor.py               108 lines   ExcelCursor
├── exceptions.py            52 lines   PEP 249 exceptions
└── engine/
    ├── base.py              43 lines   BaseEngine ABC
    ├── openpyxl_engine.py   97 lines   OpenpyxlEngine
    ├── pandas_engine.py     57 lines   PandasEngine
    ├── parser.py           365 lines   SQL parser
    ├── executor.py          31 lines   Query dispatcher
    ├── openpyxl_executor.py 283 lines   OpenpyxlExecutor
    ├── pandas_executor.py   200 lines   PandasExecutor
    └── result.py            14 lines   ExecutionResult dataclass
                           ──────
                          1,373 lines total
```

---

## 2. Detailed Module Design

### 2.1 Module: `__init__.py`

#### Purpose
Expose PEP 249-mandated module-level attributes and the `connect()` factory.

#### Design Pattern
**Factory Method** — `connect()` abstracts away `ExcelConnection` instantiation.

#### Implementation

```python
from .connection import ExcelConnection

# PEP 249 module-level constants
apilevel = "2.0"       # This module supports DB-API 2.0
threadsafety = 1       # Threads can share the module but not connections
paramstyle = "qmark"   # Uses ? for parameter placeholders

def connect(
    file_path: str,
    engine: str = "openpyxl",
    autocommit: bool = True,
    create: bool = False,
    data_only: bool = True,
) -> ExcelConnection:
    """Create a new DB-API 2.0 connection to an Excel file."""
    return ExcelConnection(
        file_path,
        engine=engine,
        autocommit=autocommit,
        create=create,
        data_only=data_only,
    )
```

#### API Extensions Beyond PEP 249

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `engine` | `str` | `"openpyxl"` | Execution engine: `"openpyxl"` or `"pandas"` |
| `autocommit` | `bool` | `True` | Auto-save write operations |
| `create` | `bool` | `False` | Create file if it doesn't exist |
| `data_only` | `bool` | `True` | Read cached formula values (openpyxl) |

### 2.2 Module: `connection.py`

#### Purpose
PEP 249 Connection implementation with engine management and transaction control.

#### Design Patterns
- **Decorator Pattern** — `check_closed` wraps methods to guard against closed connections
- **Strategy Pattern** — Engine selection at construction time
- **Context Manager** — `__enter__` / `__exit__` for `with` statement support

#### `check_closed` Decorator

```python
def check_closed(func):
    """Decorator to check if connection is closed before executing method."""
    def wrapper(self, *args, **kwargs):
        if self.closed:
            raise InterfaceError("Connection is already closed")
        return func(self, *args, **kwargs)
    return wrapper
```

Applied to: `cursor()`, `commit()`, `rollback()`

#### Engine Initialization

```python
def __init__(self, file_path, engine="openpyxl", autocommit=True, create=False, data_only=True):
    if engine == "openpyxl":
        self.engine = OpenpyxlEngine(file_path, data_only=data_only, create=create)
    elif engine == "pandas":
        self.engine = PandasEngine(file_path, data_only=data_only, create=create)
    else:
        raise InterfaceError(f"Unsupported engine: {engine}")

    self._snapshot = self.engine.snapshot()  # Initial rollback point
```

#### Transaction Methods

```python
@check_closed
def commit(self):
    self.engine.save()                        # Persist to disk
    self._snapshot = self.engine.snapshot()    # New rollback point

@check_closed
def rollback(self):
    if self.autocommit:
        raise NotSupportedError("Rollback is disabled when autocommit is enabled")
    self.engine.restore(self._snapshot)       # Revert to last commit
```

#### `workbook` Property

```python
@property
def workbook(self):
    wb = getattr(self.engine, "workbook", None)
    if wb is None:
        raise NotSupportedError(f"Engine '{self.engine_name}' does not expose a workbook object")
    return wb
```

- **OpenpyxlEngine**: Returns the `openpyxl.Workbook` object
- **PandasEngine**: Raises `NotSupportedError` (no `workbook` attribute)

### 2.3 Module: `cursor.py`

#### Purpose
PEP 249 Cursor implementation with result management and auto-save logic.

#### Design Patterns
- **Iterator Pattern** — `fetchone()` advances internal `_index` cursor
- **Template Method** — Both `execute()` and `executemany()` follow: parse → execute → handle auto-save
- **Decorator Pattern** — `check_closed` guards all public methods

#### State Management

```python
class ExcelCursor:
    def __init__(self, connection):
        self.connection = connection
        self.closed = False
        self._results: List[tuple] = []    # Result buffer
        self._index: int = 0               # Fetch cursor position
        self.description = None            # PEP 249 column metadata
        self.rowcount = -1                 # -1 before first execute
        self.lastrowid = None
        self.arraysize = 1                 # Default fetchmany size
```

#### execute() — Single Query

```python
@check_closed
def execute(self, query, params=None):
    try:
        result = self.connection.engine.execute_with_params(query, params)
    except ValueError as exc:
        raise ProgrammingError(str(exc)) from exc
    except NotImplementedError as exc:
        raise NotSupportedError(str(exc)) from exc

    self._results = result.rows
    self._index = 0
    self.description = result.description
    self.rowcount = result.rowcount
    self.lastrowid = result.lastrowid

    # Auto-save for write operations in autocommit mode
    if self.connection.autocommit and result.action in {"INSERT", "CREATE", "DROP", "UPDATE", "DELETE"}:
        self.connection.engine.save()

    return self
```

#### executemany() — Batch Execution

```python
@check_closed
def executemany(self, query, seq_of_params):
    total_rowcount = 0
    snapshot = None

    # Take snapshot for atomic batch (autocommit=False only)
    if not self.connection.autocommit:
        snapshot = self.connection.engine.snapshot()

    for params in seq_of_params:
        try:
            result = self.connection.engine.execute_with_params(query, params)
        except (ValueError, NotImplementedError) as exc:
            if snapshot is not None:
                self.connection.engine.restore(snapshot)  # Rollback entire batch
            raise ...

        total_rowcount += result.rowcount

    # Auto-save after entire batch
    if self.connection.autocommit and last_action in write_actions:
        self.connection.engine.save()
```

**Atomicity guarantee**: When `autocommit=False`, `executemany()` takes a snapshot before the loop. If any parameter set fails, the entire batch is rolled back.

#### Fetch Methods

```python
def fetchone(self):
    if self._index >= len(self._results):
        return None
    result = self._results[self._index]
    self._index += 1
    return result

def fetchall(self):
    results = self._results[self._index:]
    self._index = len(self._results)
    return results

def fetchmany(self, size=None):
    count = self.arraysize if size is None else size
    if count <= 0:
        return []
    start = self._index
    end = min(self._index + count, len(self._results))
    self._index = end
    return self._results[start:end]
```

### 2.4 Module: `exceptions.py`

#### Purpose
PEP 249 compliant exception hierarchy.

#### Hierarchy

```
Exception
├── Warning           — Important warnings (data truncation)
└── Error             — Base for all DB-API errors
    ├── InterfaceError — Bad interface usage (closed conn, bad engine)
    └── DatabaseError  — Database operation errors
        ├── DataError           — Value/type mismatch
        ├── OperationalError    — File I/O, connection issues
        ├── IntegrityError      — Referential integrity
        ├── InternalError       — Internal engine errors
        ├── ProgrammingError    — Bad SQL, unknown columns
        └── NotSupportedError   — Unsupported feature (rollback in autocommit)
```

### 2.5 Module: `engine/base.py`

#### Purpose
Abstract base class defining the engine contract.

#### Abstract Interface

| Method | Signature | Description |
|--------|-----------|-------------|
| `load()` | `() → Dict[str, Any]` | Load workbook into memory |
| `save()` | `() → None` | Persist changes to disk |
| `snapshot()` | `() → Any` | Capture current state |
| `restore()` | `(snapshot: Any) → None` | Restore from snapshot |
| `execute()` | `(query: str) → ExecutionResult` | Execute raw query |

#### Default Implementation

```python
def execute_with_params(self, query, params=None):
    return self.execute(query)  # Subclasses override for param support
```

### 2.6 Module: `engine/openpyxl_engine.py`

#### Purpose
Default engine using openpyxl for cell-level Excel access.

#### Initialization Flow

```
__init__(file_path, data_only=True, create=False)
  │
  ├── create=True AND file doesn't exist
  │   ├── Workbook() → workbook.save(file_path)  → Create empty xlsx
  │   └── self.workbook = new Workbook
  │
  └── File exists (or create=False)
      └── self.workbook = load_workbook(file_path, data_only=data_only)
  │
  └── self.data = {sheet: workbook[sheet] for sheet in workbook.sheetnames}
```

#### Atomic Save Algorithm

```python
def save(self):
    if self.workbook is None:
        raise ValueError("Workbook is not loaded")
    directory = os.path.dirname(self.file_path) or "."
    temp_file = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=directory) as handle:
            temp_file = handle.name
        self.workbook.save(temp_file)          # Write to temp
        os.replace(temp_file, self.file_path)  # Atomic rename
    finally:
        if temp_file and os.path.exists(temp_file):
            os.unlink(temp_file)               # Cleanup on failure
```

**Why atomic**: `os.replace()` is atomic on POSIX when source and destination are on the same filesystem. The temp file is created in the same directory to guarantee this.

#### Snapshot/Restore

```python
def snapshot(self) -> BytesIO:
    buffer = BytesIO()
    self.workbook.save(buffer)   # Serialize entire workbook
    buffer.seek(0)
    return buffer

def restore(self, snapshot: BytesIO):
    snapshot.seek(0)
    self.workbook = load_workbook(snapshot, data_only=self._data_only)
    self.data = {sheet: self.workbook[sheet] for sheet in self.workbook.sheetnames}
```

### 2.7 Module: `engine/pandas_engine.py`

#### Purpose
Alternative engine using pandas DataFrames for in-memory data.

#### Key Differences from OpenpyxlEngine

| Aspect | OpenpyxlEngine | PandasEngine |
|--------|----------------|-------------|
| In-memory format | openpyxl Worksheet objects | pandas DataFrames |
| Executor | OpenpyxlExecutor (via executor.py) | PandasExecutor (direct) |
| Snapshot | BytesIO serialization | Deep copy of DataFrames |
| Save | Workbook.save() preserves formatting | ExcelWriter rewrites — drops formatting |
| `workbook` property | Returns openpyxl Workbook | N/A (raises NotSupportedError) |
| File creation | `Workbook()` + save | openpyxl `Workbook()` + save + `pd.read_excel()` |

#### Snapshot/Restore (Deep Copy)

```python
def snapshot(self):
    return {name: frame.copy(deep=True) for name, frame in self.data.items()}

def restore(self, snapshot):
    self.data = {name: frame.copy(deep=True) for name, frame in snapshot.items()}
```

### 2.8 Module: `engine/parser.py`

#### Purpose
Custom recursive-descent SQL parser. No external dependencies.

#### Parser Architecture

```
parse_sql(query, params)
  │
  ├── Tokenize: query.strip().split()
  ├── Identify action: tokens[0].upper()
  │
  ├── SELECT → _parse_select(query, params)
  │   ├── Find FROM keyword (case-insensitive)
  │   ├── Extract columns (before FROM)
  │   ├── Extract table name (after FROM)
  │   ├── Parse remainder for WHERE / ORDER BY / LIMIT
  │   └── Bind parameters (? → values)
  │
  ├── INSERT → _parse_insert(query, params)
  │   ├── Find VALUES keyword
  │   ├── Extract table name and optional column list
  │   ├── Parse values inside parentheses
  │   └── Bind parameters
  │
  ├── UPDATE → _parse_update(query, params)
  │   ├── Find SET keyword
  │   ├── Extract table name
  │   ├── Parse SET assignments (col = value, ...)
  │   ├── Parse optional WHERE clause
  │   └── Bind parameters (SET values first, then WHERE values)
  │
  ├── DELETE → _parse_delete(query, params)
  │   ├── Validate DELETE FROM structure
  │   ├── Extract table name
  │   ├── Parse optional WHERE clause
  │   └── Bind parameters
  │
  ├── CREATE → _parse_create(query)
  │   ├── Validate CREATE TABLE structure
  │   ├── Extract table name and column list
  │   └── No parameter binding
  │
  └── DROP → _parse_drop(query)
      ├── Validate DROP TABLE structure
      └── Extract table name
```

#### Helper Functions

| Function | Purpose | Input → Output |
|----------|---------|----------------|
| `_parse_value(token)` | Parse a SQL literal | `"'hello'"` → `"hello"`, `"42"` → `42`, `"NULL"` → `None` |
| `_parse_columns(text)` | Parse column list | `"a, b, c"` → `["a", "b", "c"]`, `"*"` → `["*"]` |
| `_parse_where_expression(text, params)` | Parse WHERE conditions | `"id > 10 AND name = 'A'"` → `{conditions: [...], conjunctions: ["AND"]}` |
| `_bind_params(values, params)` | Replace `?` with params | `["?", "hello"]` + `(42,)` → `[42, "hello"]` |
| `_split_csv(text)` | Split by commas, respecting quotes | `"a, 'b,c', d"` → `["a", "'b,c'", "d"]` |

### 2.9 Module: `engine/executor.py`

#### Purpose
Query dispatch entry point for OpenpyxlEngine. Performs case-insensitive table lookup.

#### Algorithm

```python
def execute_query(parsed, data, workbook):
    action = parsed["action"]
    table = parsed["table"].lower()
    data_lower = {sheet.lower(): sheet for sheet in data.keys()}

    # Validate table exists (except CREATE/DROP which modify sheets)
    if action in {"SELECT", "INSERT", "UPDATE", "DELETE"}:
        if table not in data_lower:
            raise ValueError(f"Sheet '{table}' not found in Excel")

    return OpenpyxlExecutor(data, workbook).execute(parsed)
```

**Note**: The PandasEngine does NOT use this dispatcher. It directly instantiates `PandasExecutor`.

### 2.10 Module: `engine/openpyxl_executor.py`

#### Purpose
Execute parsed SQL queries using openpyxl Worksheet objects.

#### SELECT Algorithm

```
1. Get worksheet: ws = data[table]
2. Read all rows: rows = list(ws.iter_rows(values_only=True))
3. Headers = rows[0], Data = rows[1:]
4. Convert to list of dicts: [{header: value, ...}, ...]
5. Apply WHERE filter: _matches_where(row_dict, where)
6. Apply ORDER BY: sorted(data, key=_sort_key(row[col]), reverse=DESC)
7. Apply LIMIT: data[:limit]
8. Project columns: tuple(row[col] for col in selected_columns)
9. Build description: [(col, None, None, None, None, None, None), ...]
10. Return ExecutionResult
```

#### INSERT Algorithm

```
1. Get worksheet: ws = data[table]
2. Headers = row 1 values (required — empty sheet raises ValueError)
3. If column-specified: map values to header positions, fill gaps with None
4. If column-unspecified: verify len(values) == len(headers)
5. ws.append(row_values)
6. lastrowid = ws.max_row
```

#### UPDATE Algorithm

```
1. Get worksheet: ws = data[table]
2. Validate SET columns exist in headers
3. For each row (row_index 2 to max_row):
   a. Build row_dict from cells
   b. If WHERE matches (or no WHERE):
      - For each SET assignment: ws.cell(row, col, value=new_value)
      - Increment rowcount
```

#### DELETE Algorithm (Reverse Iteration)

```
1. Get worksheet: ws = data[table]
2. For each row (max_row DOWN TO 2):
   a. Build row_dict from cells
   b. If WHERE matches (or no WHERE):
      - ws.delete_rows(row_index)
      - Increment rowcount
```

**Why reverse?** Deleting row N shifts rows N+1, N+2, ... up by one. Forward iteration would skip rows after a delete.

#### WHERE Evaluation

```python
def _matches_where(self, row, where):
    # Evaluate first condition
    result = _evaluate_condition(row, conditions[0])
    # Apply conjunctions left-to-right (no precedence, no grouping)
    for idx, conj in enumerate(conjunctions):
        next_result = _evaluate_condition(row, conditions[idx + 1])
        if conj == "AND":
            result = result and next_result
        else:  # OR
            result = result or next_result
    return result

def _evaluate_condition(self, row, condition):
    left, right = _coerce_for_compare(row[column], value)
    # Operators: =, ==, !=, <>, >, >=, <, <=
```

#### Type Coercion

```python
def _coerce_for_compare(self, left, right):
    left_num = _to_number(left)
    right_num = _to_number(right)
    if left_num is not None and right_num is not None:
        return left_num, right_num   # Numeric comparison
    return str(left), str(right)     # String comparison fallback

def _to_number(self, value):
    if isinstance(value, bool): return None
    if isinstance(value, (int, float)): return float(value)
    if isinstance(value, str):
        try: return float(value)
        except ValueError: return None
    return None
```

### 2.11 Module: `engine/pandas_executor.py`

#### Purpose
Execute parsed SQL queries using pandas DataFrame operations.

#### Key Differences from OpenpyxlExecutor

| Aspect | OpenpyxlExecutor | PandasExecutor |
|--------|-----------------|----------------|
| SELECT filter | Python loop + dict matching | `frame[mask]` (vectorized) |
| SELECT sort | `sorted(data, key=...)` | `frame.sort_values()` |
| INSERT | `ws.append()` | `pd.concat([frame, new_row])` |
| UPDATE | `ws.cell(row, col, value=)` | `frame.loc[mask, col] = value` |
| DELETE | `ws.delete_rows()` (reverse) | `frame.loc[~mask].reset_index()` |
| CREATE | `workbook.create_sheet()` | `pd.DataFrame(columns=cols)` |
| WHERE | Python condition evaluation | `pd.Series` boolean mask |

#### Boolean Mask Construction

```python
def _build_mask(self, frame, where):
    conditions = where["conditions"]
    conjunctions = where["conjunctions"]
    mask = _evaluate_condition(frame, conditions[0])  # pd.Series[bool]
    for idx, conj in enumerate(conjunctions):
        next_mask = _evaluate_condition(frame, conditions[idx + 1])
        if conj == "AND":
            mask = mask & next_mask    # Bitwise AND
        else:
            mask = mask | next_mask    # Bitwise OR
    return mask
```

### 2.12 Module: `engine/result.py`

#### Purpose
Standardized query result container shared across all executors.

```python
Description = Sequence[Tuple[
    Optional[str],   # name
    Optional[str],   # type_code
    Optional[int],   # display_size
    Optional[int],   # internal_size
    Optional[int],   # precision
    Optional[int],   # scale
    Optional[bool],  # null_ok
]]

@dataclass
class ExecutionResult:
    action: str                     # SQL action type
    rows: List[Tuple]              # Result rows (empty for DML)
    description: Description        # Column metadata
    rowcount: int                   # Rows affected/returned
    lastrowid: Optional[int] = None # Last inserted row index
```

---

## 3. SQL Grammar Specification

### 3.1 Supported Grammar (BNF-style)

```bnf
<statement>    ::= <select> | <insert> | <update> | <delete> | <create> | <drop>

<select>       ::= SELECT <columns> FROM <table> [<where>] [<order_by>] [<limit>]
<columns>      ::= "*" | <column_name> ("," <column_name>)*
<table>        ::= <identifier>

<insert>       ::= INSERT INTO <table> [<col_list>] VALUES "(" <value_list> ")"
<col_list>     ::= "(" <column_name> ("," <column_name>)* ")"
<value_list>   ::= <value> ("," <value>)*

<update>       ::= UPDATE <table> SET <assignments> [<where>]
<assignments>  ::= <assignment> ("," <assignment>)*
<assignment>   ::= <column_name> "=" <value>

<delete>       ::= DELETE FROM <table> [<where>]

<create>       ::= CREATE TABLE <table> "(" <column_name> ("," <column_name>)* ")"
<drop>         ::= DROP TABLE <table>

<where>        ::= WHERE <condition> (("AND" | "OR") <condition>)*
<condition>    ::= <column_name> <operator> <value>
<operator>     ::= "=" | "==" | "!=" | "<>" | ">" | ">=" | "<" | "<="

<order_by>     ::= ORDER BY <column_name> ["ASC" | "DESC"]
<limit>        ::= LIMIT <integer>

<value>        ::= <string> | <integer> | <float> | NULL | "?"
<string>       ::= "'" <text> "'" | '"' <text> '"'
<integer>      ::= ["-"] <digit>+
<float>        ::= ["-"] <digit>+ "." <digit>+
<identifier>   ::= <letter_or_digit>+    (NO QUOTES — unquoted only)
```

### 3.2 Parameter Binding Rules

1. **Paramstyle**: `qmark` — uses `?` as placeholder
2. **Binding order** (left-to-right in SQL):
   - SELECT: WHERE values → LIMIT value
   - INSERT: VALUES list
   - UPDATE: SET values → WHERE values
   - DELETE: WHERE values
3. **Validation**:
   - `?` without params → `ValueError("Missing parameters")`
   - More params than `?` → `ValueError("Too many parameters")`
   - Fewer params than `?` → `ValueError("Not enough parameters")`

### 3.3 Limitations

| Feature | Status | Notes |
|---------|--------|-------|
| Subqueries | ❌ | Not supported |
| JOIN | ❌ | Not supported |
| GROUP BY / HAVING | ❌ | Not supported |
| Aggregate functions | ❌ | COUNT, SUM, etc. not supported |
| Parenthesized WHERE | ❌ | `(A AND B) OR C` not supported |
| LIKE / IN / BETWEEN | ❌ | Not supported |
| DISTINCT | ❌ | Not supported |
| Aliases | ❌ | `AS` not supported |
| Quoted table names | ❌ | `"Sheet1"` causes lookup failure — use `Sheet1` |
| Multiple ORDER BY | ❌ | Only single column sort |
| OFFSET | ❌ | Only LIMIT, no OFFSET |

---

## 4. Engine Abstraction Patterns

### 4.1 Strategy Pattern

The engine is selected at connection time and encapsulated behind the `BaseEngine` interface:

```python
# Connection delegates all operations to the engine
class ExcelConnection:
    def __init__(self, file_path, engine="openpyxl", ...):
        if engine == "openpyxl":
            self.engine = OpenpyxlEngine(file_path, ...)
        elif engine == "pandas":
            self.engine = PandasEngine(file_path, ...)

    # All operations go through self.engine
    def commit(self):  self.engine.save(); ...
    def rollback(self): self.engine.restore(self._snapshot)
```

### 4.2 Engine Parity

Both engines implement the same interface and produce equivalent results for the same queries:

```
OpenpyxlEngine.execute("SELECT * FROM Sheet1 WHERE id > 10")
  → ExecutionResult(rows=[...], description=[...], rowcount=N)

PandasEngine.execute("SELECT * FROM Sheet1 WHERE id > 10")
  → ExecutionResult(rows=[...], description=[...], rowcount=N)
  # Same rows, same description format, same rowcount
```

### 4.3 Engine-Specific Behaviors

| Behavior | OpenpyxlEngine | PandasEngine |
|----------|----------------|-------------|
| `None` handling in WHERE | `str(None)` = `"None"` for string compare | pandas `NaN` behavior |
| Numeric types | Returns original Excel types | Returns pandas types (may convert) |
| Row ordering | Preserves original Excel row order | Preserves DataFrame index order |
| Save side effects | Preserves formatting, charts | Drops formatting, charts, formulas |

---

## 5. Configuration: `pyproject.toml`

```toml
[build-system]
requires = ["setuptools>=61.0"]
build-backend = "setuptools.build_meta"

[project]
name = "excel-dbapi"
version = "1.0.0"
description = "PEP 249 compliant DB-API driver for Excel files"
readme = "README.md"
license = { text = "MIT" }
requires-python = ">=3.10"
dependencies = [
    "pandas>=2.0.0,<3.0.0",
    "openpyxl>=3.1.0,<4.0.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.3.5,<9.0.0",
    "pytest-cov>=4.1,<5.0.0",
    "black>=25.1.0,<26.0.0",
    "isort>=6.0.1,<7.0.0",
    "ruff>=0.11.2,<0.12.0",
    "mypy>=1.15.0,<2.0.0",
    "build",
    "bandit>=1.7.0,<2.0.0",
    "vulture>=2.0,<3.0",
    "pre-commit",
    "types-requests>=2.28.0,<3.0.0",
]

[tool.mypy]
ignore_missing_imports = true
strict = false

[tool.pytest.ini_options]
pythonpath = ["src"]

[tool.setuptools.packages.find]
where = ["src"]
```

### Dependency Rationale

| Dependency | Why Required |
|-----------|-------------|
| `pandas>=2.0.0` | PandasEngine, DataFrame-based queries |
| `openpyxl>=3.1.0` | OpenpyxlEngine (default), also used by PandasEngine for file creation |

---

## 6. API Usage Examples

### 6.1 Basic SELECT Query

```python
from excel_dbapi import connect

conn = connect("employees.xlsx")
cursor = conn.cursor()

cursor.execute("SELECT * FROM Sheet1")
print(cursor.description)  # [('id', None, ...), ('name', None, ...), ...]
print(cursor.rowcount)     # Number of rows returned

for row in cursor.fetchall():
    print(row)  # (1, 'Alice', 'Engineering')

conn.close()
```

### 6.2 Filtered Query with Parameters

```python
from excel_dbapi import connect

with connect("sales.xlsx") as conn:
    cursor = conn.cursor()

    # Parameter binding with ?
    cursor.execute(
        "SELECT product, revenue FROM Sales WHERE region = ? AND revenue > ?",
        ("APAC", 10000)
    )

    for row in cursor.fetchall():
        print(f"Product: {row[0]}, Revenue: {row[1]}")
```

### 6.3 Sorting and Limiting Results

```python
with connect("data.xlsx") as conn:
    cursor = conn.cursor()

    cursor.execute(
        "SELECT name, score FROM Students ORDER BY score DESC LIMIT 10"
    )

    print("Top 10 students:")
    for rank, row in enumerate(cursor.fetchall(), 1):
        print(f"  {rank}. {row[0]} — {row[1]}")
```

### 6.4 Inserting Data

```python
with connect("inventory.xlsx") as conn:
    cursor = conn.cursor()

    # Column-specified insert
    cursor.execute(
        "INSERT INTO Products (id, name, price) VALUES (?, ?, ?)",
        (101, "Widget", 9.99)
    )
    print(f"Inserted at row: {cursor.lastrowid}")

    # Bulk insert with executemany
    items = [
        (102, "Gadget", 19.99),
        (103, "Doohickey", 4.99),
        (104, "Thingamajig", 29.99),
    ]
    cursor.executemany(
        "INSERT INTO Products (id, name, price) VALUES (?, ?, ?)",
        items
    )
    print(f"Inserted {cursor.rowcount} rows")
```

### 6.5 Updating Data

```python
with connect("data.xlsx") as conn:
    cursor = conn.cursor()

    cursor.execute(
        "UPDATE Employees SET department = ? WHERE department = ?",
        ("Engineering", "Eng")
    )
    print(f"Updated {cursor.rowcount} rows")
```

### 6.6 Deleting Data

```python
with connect("data.xlsx") as conn:
    cursor = conn.cursor()

    # Conditional delete
    cursor.execute("DELETE FROM Logs WHERE date < ?", ("2025-01-01",))
    print(f"Deleted {cursor.rowcount} rows")

    # Delete all rows (keep headers)
    cursor.execute("DELETE FROM TempData")
    print(f"Cleared {cursor.rowcount} rows from TempData")
```

### 6.7 Creating and Dropping Sheets

```python
with connect("workbook.xlsx") as conn:
    cursor = conn.cursor()

    # Create a new sheet with headers
    cursor.execute("CREATE TABLE Reports (id, title, date, status)")

    # Insert data into the new sheet
    cursor.execute(
        "INSERT INTO Reports (id, title, date, status) VALUES (?, ?, ?, ?)",
        (1, "Q1 Report", "2026-03-15", "Draft")
    )

    # Drop a sheet
    cursor.execute("DROP TABLE OldSheet")
```

### 6.8 Transactions (Manual Commit/Rollback)

```python
from excel_dbapi import connect

conn = connect("ledger.xlsx", autocommit=False)
cursor = conn.cursor()

try:
    cursor.execute("UPDATE Accounts SET balance = balance - 100 WHERE id = ?", (1,))
    cursor.execute("UPDATE Accounts SET balance = balance + 100 WHERE id = ?", (2,))

    # Verify the transfer
    cursor.execute("SELECT balance FROM Accounts WHERE id = ?", (1,))
    balance = cursor.fetchone()[0]

    if balance < 0:
        conn.rollback()   # Undo both updates
        print("Transfer failed: insufficient funds")
    else:
        conn.commit()     # Persist both updates
        print("Transfer completed")

except Exception:
    conn.rollback()
finally:
    conn.close()
```

### 6.9 Using the Pandas Engine

```python
from excel_dbapi import connect

with connect("data.xlsx", engine="pandas") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1 WHERE value > ?", (100,))

    for row in cursor.fetchall():
        print(row)
```

### 6.10 Creating a New File

```python
from excel_dbapi import connect

# Create a new Excel file from scratch
conn = connect("new_file.xlsx", create=True)
cursor = conn.cursor()

cursor.execute("CREATE TABLE Users (id, name, email)")
cursor.execute(
    "INSERT INTO Users (id, name, email) VALUES (?, ?, ?)",
    (1, "Alice", "alice@example.com")
)

conn.close()
```

### 6.11 Accessing the Workbook Directly

```python
from excel_dbapi import connect

with connect("report.xlsx") as conn:
    # SQL channel: query data
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Data")

    # Workbook channel: format cells
    wb = conn.workbook
    ws = wb["Data"]
    ws["A1"].font = ws["A1"].font.copy(bold=True)
    ws.column_dimensions["A"].width = 20

    conn.commit()  # Save both data and formatting changes
```

### 6.12 Fetch Methods

```python
with connect("data.xlsx") as conn:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Sheet1")

    # Fetch one row at a time
    first_row = cursor.fetchone()   # Returns tuple or None
    second_row = cursor.fetchone()

    # Fetch N rows at a time
    cursor.arraysize = 5
    batch = cursor.fetchmany()  # Returns up to 5 rows
    batch = cursor.fetchmany(10)  # Returns up to 10 rows

    # Fetch all remaining
    rest = cursor.fetchall()  # Returns list of tuples
```

---

## 7. Testing Strategy

### 7.1 Test Summary

| Category | Count | Description |
|----------|-------|-------------|
| Connection tests | ~15 | Open/close, engine selection, context manager |
| Cursor tests | ~20 | Execute, fetch methods, metadata attributes |
| Parser tests | ~25 | All SQL statements, edge cases, error handling |
| OpenpyxlExecutor tests | ~10 | All CRUD + DDL operations |
| PandasExecutor tests | ~10 | Engine parity tests |
| Transaction tests | ~5 | Commit, rollback, snapshot lifecycle |
| **Total** | **85** | All passing ✅ |

### 7.2 Test Categories

#### Connection Tests
```python
def test_connection_open_close():
    conn = ExcelConnection("test.xlsx")
    assert not conn.closed
    conn.close()
    assert conn.closed

def test_connection_closed_error():
    conn = ExcelConnection("test.xlsx")
    conn.close()
    with pytest.raises(InterfaceError):
        conn.cursor()

def test_unsupported_engine():
    with pytest.raises(InterfaceError):
        ExcelConnection("test.xlsx", engine="invalid")
```

#### Parser Tests
```python
def test_parse_select_basic():
    result = parse_sql("SELECT * FROM Sheet1")
    assert result["action"] == "SELECT"
    assert result["table"] == "Sheet1"
    assert result["columns"] == ["*"]

def test_parse_select_with_where():
    result = parse_sql("SELECT name FROM S WHERE id = ?", (42,))
    assert result["where"]["conditions"][0]["value"] == 42

def test_parse_invalid_sql():
    with pytest.raises(ValueError):
        parse_sql("INVALID QUERY")
```

#### Engine Parity Tests
```python
@pytest.mark.parametrize("engine_name", ["openpyxl", "pandas"])
def test_select_both_engines(engine_name, sample_xlsx):
    with ExcelConnection(sample_xlsx, engine=engine_name) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        rows = cursor.fetchall()
        assert len(rows) > 0
```

### 7.3 CI Configuration

```yaml
# .github/workflows/ci.yml
name: CI
on: [push, pull_request]
jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        python-version: ["3.10", "3.11", "3.12", "3.13"]
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: ${{ matrix.python-version }}
      - run: pip install -e ".[dev]"
      - run: pytest --cov=excel_dbapi
```

---

## 8. Interoperability with sqlalchemy-excel

### 8.1 Integration Points

sqlalchemy-excel uses excel-dbapi as its **full dependency** for all Excel I/O:

```python
# sqlalchemy-excel's ExcelWorkbookSession wraps excel-dbapi
from excel_dbapi import connect

class ExcelWorkbookSession:
    def __init__(self, path):
        self._conn = connect(path, engine="openpyxl", create=True)

    @property
    def workbook(self):
        return self._conn.workbook  # openpyxl Workbook

    def execute(self, sql, params=None):
        cursor = self._conn.cursor()
        cursor.execute(sql, params)
        return cursor

    def close(self):
        self._conn.commit()
        self._conn.close()
```

```python
# sqlalchemy-excel's ExcelDbapiReader uses excel-dbapi for SQL reads
class ExcelDbapiReader:
    def read(self, conn, sheet_name):
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM {sheet_name}")
        headers = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        return headers, [dict(zip(headers, row)) for row in rows]
```

### 8.2 API Contract

excel-dbapi guarantees these interfaces for sqlalchemy-excel:

| Interface | Guarantee |
|-----------|-----------|
| `connect(path, engine="openpyxl", create=True)` | Returns valid ExcelConnection |
| `conn.workbook` | Returns openpyxl Workbook (openpyxl engine) |
| `conn.cursor()` | Returns PEP 249 Cursor |
| `cursor.execute("SELECT * FROM Sheet")` | Populates `description` and result set |
| `cursor.description` | 7-tuple format per PEP 249 |
| `cursor.fetchall()` | Returns `List[Tuple]` |
| `create=True` | Creates valid empty workbook if file missing |
| All exceptions | Follow PEP 249 hierarchy |

### 8.3 Dependency Specification

```toml
# In sqlalchemy-excel's pyproject.toml
[project]
dependencies = [
    "excel-dbapi>=1.0",
    ...
]
```

---

## 9. Known Issues and Workarounds

| Issue | Workaround |
|-------|-----------|
| Quoted table names fail | Use unquoted: `Sheet1` not `"Sheet1"` |
| PandasEngine drops formatting | Use OpenpyxlEngine for formatting-sensitive workbooks |
| WHERE has no parenthesis grouping | Chain conditions carefully; `A AND B OR C` evaluates left-to-right |
| Large file memory usage | No streaming mode — full workbook loaded into memory |
| `rollback()` in autocommit | Disable autocommit: `connect(path, autocommit=False)` |
| Multiple ORDER BY columns | Not supported — sort by one column only |
| Formula cells | Set `data_only=True` (default) to read cached values |

---

## 10. Future Technical Design (v2.0.x)

### 10.1 SQLAlchemy Dialect

```python
# Planned implementation
class ExcelDialect(DefaultDialect):
    name = "excel"
    driver = "excel_dbapi"

    @classmethod
    def dbapi(cls):
        import excel_dbapi
        return excel_dbapi

    def create_connect_args(self, url):
        return [], {"file_path": url.database}
```

### 10.2 JOIN Support (Planned)

```sql
-- Planned syntax
SELECT a.name, b.department
FROM Employees a JOIN Departments b ON a.dept_id = b.id
WHERE b.name = 'Engineering'
```

### 10.3 Aggregate Functions (Planned)

```sql
-- Planned syntax
SELECT department, COUNT(*), AVG(salary)
FROM Employees
GROUP BY department
HAVING COUNT(*) > 5
```
