# ARCH.md — excel-dbapi Architecture Document

> **Version**: 1.0.0  
> **Last Updated**: 2026-03-15  
> **Status**: v1.0.0 Stable (feature-complete)

---

## 1. System Overview

excel-dbapi provides a PEP 249 (DB-API 2.0) compliant interface to read and write Excel (.xlsx) files using SQL. The system is organized in four layers: **Interface**, **Parsing**, **Execution**, and **Storage**.

```
┌─────────────────────────────────────────────────────────────────┐
│                      User Application                           │
│  conn = connect("data.xlsx")                                    │
│  cursor = conn.cursor()                                         │
│  cursor.execute("SELECT name FROM Sheet1 WHERE id > 10")       │
│  rows = cursor.fetchall()                                       │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                               ▼
┌──────────────────────────────────────────────────────────────────┐
│  Layer 1: DB-API 2.0 Interface                                   │
│  ┌──────────────────┐    ┌───────────────────┐                   │
│  │  ExcelConnection │───▶│   ExcelCursor     │                   │
│  │  • cursor()      │    │   • execute()     │                   │
│  │  • commit()      │    │   • executemany() │                   │
│  │  • rollback()    │    │   • fetchone()    │                   │
│  │  • close()       │    │   • fetchall()    │                   │
│  │  • workbook      │    │   • fetchmany()   │                   │
│  └──────────────────┘    │   • description   │                   │
│                          │   • rowcount      │                   │
│                          └─────────┬─────────┘                   │
│  ┌──────────────────────────────┐  │                             │
│  │  exceptions.py              │  │                             │
│  │  PEP 249 Exception Hierarchy│  │                             │
│  └──────────────────────────────┘  │                             │
└────────────────────────────────────┼─────────────────────────────┘
                                     │  execute_with_params(query, params)
                                     ▼
┌──────────────────────────────────────────────────────────────────┐
│  Layer 2: SQL Parsing                                            │
│  ┌──────────────────────────────────────────────────────────┐    │
│  │  parser.py — Custom Recursive-Descent SQL Parser         │    │
│  │                                                          │    │
│  │  parse_sql(query, params) → Dict[str, Any]              │    │
│  │                                                          │    │
│  │  Supported:                                              │    │
│  │  • _parse_select()  — SELECT ... FROM ... WHERE/ORDER/LIMIT │ │
│  │  • _parse_insert()  — INSERT INTO ... VALUES (...)       │    │
│  │  • _parse_update()  — UPDATE ... SET ... WHERE ...       │    │
│  │  • _parse_delete()  — DELETE FROM ... WHERE ...          │    │
│  │  • _parse_create()  — CREATE TABLE ... (col1, col2)      │    │
│  │  • _parse_drop()    — DROP TABLE ...                     │    │
│  │                                                          │    │
│  │  Helpers:                                                │    │
│  │  • _parse_value()   — NULL / string / int / float        │    │
│  │  • _parse_columns() — Column list or *                   │    │
│  │  • _parse_where_expression() — Conditions + conjunctions │    │
│  │  • _bind_params()   — ? placeholder substitution         │    │
│  │  • _split_csv()     — Quote-aware comma splitting        │    │
│  └──────────────────────────────────────────────────────────┘    │
└────────────────────────────────────┬─────────────────────────────┘
                                     │  parsed: Dict[str, Any]
                                     ▼
┌──────────────────────────────────────────────────────────────────┐
│  Layer 3: Engine + Execution                                     │
│                                                                  │
│  ┌─────────────────────────────────────────┐                     │
│  │  BaseEngine (ABC)                       │                     │
│  │  • load() → Dict[str, Any]             │                     │
│  │  • save()                               │                     │
│  │  • snapshot() → Any                     │                     │
│  │  • restore(snapshot)                    │                     │
│  │  • execute(query) → ExecutionResult     │                     │
│  │  • execute_with_params(query, params)   │                     │
│  └───────────────┬─────────────────────────┘                     │
│                  │                                               │
│        ┌─────────┴──────────┐                                    │
│        ▼                    ▼                                    │
│  ┌──────────────────┐ ┌──────────────────┐                      │
│  │ OpenpyxlEngine   │ │ PandasEngine     │                      │
│  │                  │ │                  │                      │
│  │ • Workbook obj   │ │ • DataFrame dict │                      │
│  │ • BytesIO snap   │ │ • deepcopy snap  │                      │
│  │ • Atomic save    │ │ • ExcelWriter    │                      │
│  │ • workbook prop  │ │ • No workbook    │                      │
│  └────────┬─────────┘ └────────┬─────────┘                      │
│           │ execute()          │ execute()                       │
│           ▼                    ▼                                 │
│  ┌──────────────────┐ ┌──────────────────┐                      │
│  │OpenpyxlExecutor  │ │ PandasExecutor   │                      │
│  │                  │ │                  │                      │
│  │ Cell-level ops   │ │ DataFrame ops    │                      │
│  │ ws.iter_rows()   │ │ pd.DataFrame     │                      │
│  │ ws.cell()        │ │ .loc[], .concat()│                      │
│  │ ws.append()      │ │ .sort_values()   │                      │
│  │ ws.delete_rows() │ │ boolean masks    │                      │
│  └──────────────────┘ └──────────────────┘                      │
│                                                                  │
│  ┌──────────────────────────────────────────┐                    │
│  │  ExecutionResult (dataclass)             │                    │
│  │  • action: str                          │                    │
│  │  • rows: List[Tuple]                    │                    │
│  │  • description: Sequence[7-tuple]       │                    │
│  │  • rowcount: int                        │                    │
│  │  • lastrowid: Optional[int]             │                    │
│  └──────────────────────────────────────────┘                    │
└────────────────────────────────────┬─────────────────────────────┘
                                     │
                                     ▼
┌──────────────────────────────────────────────────────────────────┐
│  Layer 4: Storage (Excel File)                                   │
│                                                                  │
│  ┌────────────────────────────────────────────┐                  │
│  │  .xlsx File (Excel 2007+ Open XML Format)  │                  │
│  │                                            │                  │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐   │                  │
│  │  │ Sheet1   │ │ Sheet2   │ │ Sheet3   │   │                  │
│  │  │ (Table)  │ │ (Table)  │ │ (Table)  │   │                  │
│  │  │          │ │          │ │          │   │                  │
│  │  │ Row 1:   │ │          │ │          │   │                  │
│  │  │ Headers  │ │          │ │          │   │                  │
│  │  │ Row 2+:  │ │          │ │          │   │                  │
│  │  │ Data     │ │          │ │          │   │                  │
│  │  └──────────┘ └──────────┘ └──────────┘   │                  │
│  └────────────────────────────────────────────┘                  │
└──────────────────────────────────────────────────────────────────┘
```

---

## 2. Layer Diagram

```
┌─────────────────────────────────────────────────┐
│              Application Code                    │
│  (SQL queries, DB-API method calls)             │
├─────────────────────────────────────────────────┤
│         DB-API 2.0 Interface Layer              │
│  connection.py │ cursor.py │ exceptions.py      │
├─────────────────────────────────────────────────┤
│              SQL Parsing Layer                   │
│  parser.py (parse_sql → Dict)                   │
├─────────────────────────────────────────────────┤
│         Engine Abstraction Layer                 │
│  base.py (BaseEngine ABC)                       │
│  executor.py (dispatch + table lookup)          │
├────────────────────┬────────────────────────────┤
│  OpenpyxlEngine    │  PandasEngine              │
│  openpyxl_engine   │  pandas_engine             │
│  OpenpyxlExecutor  │  PandasExecutor            │
├────────────────────┴────────────────────────────┤
│              Storage Layer                       │
│  .xlsx file (openpyxl / pandas I/O)             │
└─────────────────────────────────────────────────┘
```

---

## 3. Module Specifications

### 3.1 `__init__.py` — Module Entry Point

**Responsibility**: Expose PEP 249 module-level constants and the `connect()` factory function.

```python
# Module-level DB-API 2.0 constants
apilevel = "2.0"       # DB-API version
threadsafety = 1       # Threads may share module, not connections
paramstyle = "qmark"   # ? placeholder for parameters

# Factory function
def connect(file_path, engine="openpyxl", autocommit=True, create=False, data_only=True):
    return ExcelConnection(file_path, engine=engine, autocommit=autocommit,
                           create=create, data_only=data_only)
```

**Exports**: `ExcelConnection`, `connect`, `apilevel`, `threadsafety`, `paramstyle`

### 3.2 `connection.py` — ExcelConnection

**Responsibility**: PEP 249 Connection object. Manages engine lifecycle, cursor creation, and transaction control.

```
ExcelConnection
├── __init__(file_path, engine, autocommit, create, data_only)
│   ├── Instantiates OpenpyxlEngine or PandasEngine
│   └── Takes initial snapshot for rollback
├── cursor() → ExcelCursor          [check_closed]
├── commit() → None                 [check_closed]
│   ├── engine.save()               — Persist to disk
│   └── engine.snapshot()           — New rollback point
├── rollback() → None               [check_closed]
│   ├── Raises NotSupportedError if autocommit=True
│   └── engine.restore(snapshot)    — Revert to last commit
├── close() → None                  — Sets closed=True
├── workbook (property)             — engine.workbook or NotSupportedError
├── engine_name (property)          — Engine class name
├── __enter__ / __exit__            — Context manager (auto-close)
└── __str__ / __repr__              — Debug representation
```

**Key Pattern**: `check_closed` decorator wraps `cursor()`, `commit()`, `rollback()` to raise `InterfaceError` on closed connections.

### 3.3 `cursor.py` — ExcelCursor

**Responsibility**: PEP 249 Cursor object. Executes queries via the engine and manages result iteration.

```
ExcelCursor
├── __init__(connection)
│   ├── _results: List[tuple] = []   — In-memory result set
│   ├── _index: int = 0             — Fetch cursor position
│   ├── description = None          — PEP 249 column metadata
│   ├── rowcount = -1               — Rows affected / returned
│   ├── lastrowid = None            — Last inserted row index
│   └── arraysize = 1               — Default fetchmany size
├── execute(query, params=None) → self  [check_closed]
│   ├── engine.execute_with_params(query, params)
│   ├── Auto-save if autocommit + write action
│   └── Wraps ValueError → ProgrammingError
├── executemany(query, seq_of_params) → self  [check_closed]
│   ├── Snapshot before batch (if autocommit=False)
│   ├── Execute each param tuple
│   ├── Rollback all on any failure
│   └── Auto-save if autocommit + write action
├── fetchone() → tuple | None       [check_closed]
├── fetchall() → List[tuple]        [check_closed]
├── fetchmany(size=None) → List[tuple]  [check_closed]
└── close() → None
```

**Auto-save logic**: After `execute()` / `executemany()`, if `connection.autocommit=True` and action is a write operation (INSERT/UPDATE/DELETE/CREATE/DROP), `engine.save()` is called automatically.

### 3.4 `exceptions.py` — PEP 249 Exception Hierarchy

```
Exception
├── Warning                     — Data truncation, important warnings
└── Error                       — Base for all DB-API errors
    ├── InterfaceError           — Connection/cursor closed, bad interface usage
    └── DatabaseError            — All database-related errors
        ├── DataError            — Value out of range, type mismatch
        ├── OperationalError     — File I/O errors, connection issues
        ├── IntegrityError       — Relational integrity violations
        ├── InternalError        — Internal engine errors
        ├── ProgrammingError     — Bad SQL syntax, unknown columns
        └── NotSupportedError    — rollback in autocommit, workbook on PandasEngine
```

### 3.5 `engine/base.py` — BaseEngine (ABC)

**Responsibility**: Abstract base class defining the engine contract.

```python
class BaseEngine(ABC):
    def __init__(self, file_path: str, *, create: bool = False):
        self.file_path = file_path
        self.create = create

    @abstractmethod
    def load(self) -> Dict[str, Any]: ...       # Load file into memory
    @abstractmethod
    def save(self) -> None: ...                  # Persist to disk
    @abstractmethod
    def snapshot(self) -> Any: ...               # Capture state for rollback
    @abstractmethod
    def restore(self, snapshot: Any) -> None: ... # Restore from snapshot
    @abstractmethod
    def execute(self, query: str) -> ExecutionResult: ...

    def execute_with_params(self, query: str, params=None) -> ExecutionResult:
        return self.execute(query)  # Default: ignore params (overridden by subclasses)
```

### 3.6 `engine/openpyxl_engine.py` — OpenpyxlEngine

**Responsibility**: Default engine. Uses openpyxl for cell-level Excel access.

| Method | Implementation |
|--------|---------------|
| `__init__` | `load_workbook(path, data_only=True)` or `Workbook()` if `create=True` |
| `load()` | Returns `{sheet_name: Worksheet}` dict |
| `save()` | Atomic save: `tempfile.NamedTemporaryFile` → `workbook.save(tmp)` → `os.replace(tmp, path)` |
| `snapshot()` | Saves workbook to `BytesIO` buffer |
| `restore(snapshot)` | `load_workbook(BytesIO)` to reconstruct state |
| `execute(query)` | `parse_sql(query)` → `execute_query(parsed, data, workbook)` |
| `execute_with_params(query, params)` | `parse_sql(query, params)` → `execute_query(parsed, data, workbook)` |
| `workbook` (property) | Direct access to `openpyxl.Workbook` object |

**Atomic Save Sequence**:
```
1. Get directory of target file
2. Create NamedTemporaryFile in same directory (same filesystem)
3. workbook.save(temp_file_path)
4. os.replace(temp_file_path, target_file_path)  ← atomic on same filesystem
5. Finally: cleanup temp file if os.replace failed
```

### 3.7 `engine/pandas_engine.py` — PandasEngine

**Responsibility**: Alternative engine using pandas DataFrames for in-memory data.

| Method | Implementation |
|--------|---------------|
| `__init__` | Creates file via openpyxl if `create=True`, then `pd.read_excel(path, sheet_name=None)` |
| `load()` | Returns `{sheet_name: DataFrame}` dict |
| `save()` | Atomic: `pd.ExcelWriter(tmp)` → each frame `.to_excel()` → `os.replace(tmp, path)` |
| `snapshot()` | `{name: frame.copy(deep=True)}` for each sheet |
| `restore(snapshot)` | `{name: frame.copy(deep=True)}` from snapshot |
| `execute(query)` | `parse_sql(query)` → `PandasExecutor(data).execute(parsed)` |
| `execute_with_params(query, params)` | `parse_sql(query, params)` → `PandasExecutor(data).execute(parsed)` |

**No `workbook` property**: PandasEngine does not expose an openpyxl Workbook. Accessing `conn.workbook` raises `NotSupportedError`.

**Limitation**: PandasEngine rewrites the entire workbook on save via `pd.ExcelWriter`. This **drops formatting, charts, formulas, and other non-data elements**.

### 3.8 `engine/parser.py` — SQL Parser

**Responsibility**: Custom recursive-descent parser that converts SQL strings into parsed dictionaries.

**Input**: SQL query string + optional params tuple  
**Output**: Dict with action-specific keys

#### Parsed Output Formats:

**SELECT**:
```python
{
    "action": "SELECT",
    "table": "Sheet1",
    "columns": ["col1", "col2"] | ["*"],
    "where": {"conditions": [...], "conjunctions": [...]},  # optional
    "order_by": {"column": "col1", "direction": "ASC"},     # optional
    "limit": 10,                                             # optional
    "params": (value1, value2)                               # original params
}
```

**INSERT**:
```python
{
    "action": "INSERT",
    "table": "Sheet1",
    "columns": ["col1", "col2"] | None,  # None = all columns
    "values": [val1, val2],
    "params": (val1, val2)
}
```

**UPDATE**:
```python
{
    "action": "UPDATE",
    "table": "Sheet1",
    "set": [{"column": "name", "value": "Alice"}, ...],
    "where": {"conditions": [...], "conjunctions": [...]},  # optional
    "params": (val1, val2)
}
```

**DELETE**:
```python
{
    "action": "DELETE",
    "table": "Sheet1",
    "where": {"conditions": [...], "conjunctions": [...]},  # optional
    "params": (val1,)
}
```

**CREATE TABLE**:
```python
{
    "action": "CREATE",
    "table": "NewSheet",
    "columns": ["col1", "col2"],
    "params": None
}
```

**DROP TABLE**:
```python
{
    "action": "DROP",
    "table": "SheetName",
    "params": None
}
```

#### WHERE Expression Structure:
```python
{
    "conditions": [
        {"column": "id", "operator": ">", "value": 10},
        {"column": "name", "operator": "=", "value": "Alice"}
    ],
    "conjunctions": ["AND"]  # Between condition[0] and condition[1]
}
```

#### Parameter Binding:
```python
# Input:  "SELECT * FROM S WHERE id = ?" with params=(42,)
# Output: condition["value"] = 42 (? replaced by param)

# Binding order: WHERE conditions left-to-right, then LIMIT
# For UPDATE: SET values left-to-right, then WHERE conditions
```

### 3.9 `engine/executor.py` — Query Dispatcher

**Responsibility**: Entry point for openpyxl query execution. Validates table existence with case-insensitive lookup, then delegates to `OpenpyxlExecutor`.

```python
def execute_query(parsed, data, workbook) -> ExecutionResult:
    table = parsed["table"].lower()
    data_lower = {sheet.lower(): sheet for sheet in data.keys()}

    if action in {"SELECT", "INSERT", "UPDATE", "DELETE"}:
        if table not in data_lower:
            raise ValueError(f"Sheet '{table}' not found")

    return OpenpyxlExecutor(data, workbook).execute(parsed)
```

**Note**: CREATE and DROP bypass the existence check since they create/remove sheets.

### 3.10 `engine/openpyxl_executor.py` — OpenpyxlExecutor

**Responsibility**: Executes parsed queries using openpyxl Worksheet objects.

| Action | Algorithm |
|--------|-----------|
| SELECT | `iter_rows(values_only=True)` → headers from row 1 → filter by WHERE → sort by ORDER BY → limit by LIMIT → project columns |
| INSERT | Get headers from row 1 → map values to columns → `ws.append(row)` → return lastrowid |
| UPDATE | Iterate rows 2..max_row → match WHERE → `ws.cell(row, col, value=new_val)` → count affected |
| DELETE | Iterate rows max_row..2 (reverse) → match WHERE → `ws.delete_rows(row)` → count affected |
| CREATE | `workbook.create_sheet(title=table)` → `ws.append(columns)` → add to data dict |
| DROP | `workbook.remove(ws)` → `del data[table]` |

**DELETE reverse iteration**: Rows are deleted from bottom to top to avoid index shifting issues.

**WHERE evaluation**:
```
_matches_where(row_dict, where) → bool
├── Evaluate each condition with _evaluate_condition()
├── Apply conjunctions: AND (short-circuit &&), OR (short-circuit ||)
└── _coerce_for_compare(): attempt numeric comparison first, fall back to string
```

**Sort key**: `_sort_key(value)` returns `(0, numeric)` for numbers, `(0, str)` for strings, `(1, "")` for None. This ensures None sorts last.

### 3.11 `engine/pandas_executor.py` — PandasExecutor

**Responsibility**: Executes parsed queries using pandas DataFrame operations.

| Action | Algorithm |
|--------|-----------|
| SELECT | `frame[mask]` → `sort_values()` → `head(limit)` → project columns → `itertuples()` |
| INSERT | `pd.concat([frame, pd.DataFrame([row_data])])` |
| UPDATE | `frame.loc[mask, column] = value` |
| DELETE | `frame.loc[~mask].reset_index(drop=True)` |
| CREATE | `pd.DataFrame(columns=parsed["columns"])` → add to data dict |
| DROP | `del data[table]` |

**Boolean mask construction**: `_build_mask(frame, where)` creates a pandas Series mask using vectorized comparison operators, combined with `&` (AND) / `|` (OR).

### 3.12 `engine/result.py` — ExecutionResult

**Responsibility**: Standardized query result container.

```python
@dataclass
class ExecutionResult:
    action: str                    # "SELECT", "INSERT", "UPDATE", etc.
    rows: List[Tuple]             # Result rows (empty for write ops)
    description: Description       # PEP 249 column metadata (7-tuples)
    rowcount: int                  # Rows affected or returned
    lastrowid: Optional[int] = None  # Last inserted row index

# Description type:
Description = Sequence[Tuple[
    Optional[str],   # name
    Optional[str],   # type_code
    Optional[int],   # display_size
    Optional[int],   # internal_size
    Optional[int],   # precision
    Optional[int],   # scale
    Optional[bool],  # null_ok
]]
```

---

## 4. Data Flow Diagrams

### 4.1 Query Execution Flow

```
User: cursor.execute("SELECT name FROM Sheet1 WHERE id > 10", (10,))
  │
  ▼
ExcelCursor.execute(query, params)
  │
  ├── connection.engine.execute_with_params(query, params)
  │     │
  │     ▼
  │   parse_sql(query, params)
  │     │
  │     ├── Tokenize query → identify action (SELECT)
  │     ├── _parse_select() → extract table, columns, where, order_by, limit
  │     ├── _parse_where_expression() → conditions + conjunctions
  │     └── _bind_params() → replace ? with param values
  │     │
  │     ▼
  │   parsed = {action: "SELECT", table: "Sheet1", columns: ["name"],
  │             where: {conditions: [{column: "id", operator: ">", value: 10}]},
  │             order_by: None, limit: None}
  │     │
  │     ▼
  │   [OpenpyxlEngine path]          [PandasEngine path]
  │   execute_query(parsed, data, wb)  PandasExecutor(data).execute(parsed)
  │     │                               │
  │     ▼                               ▼
  │   OpenpyxlExecutor.execute()       DataFrame filtering via boolean mask
  │     │                               │
  │     ├── Get worksheet              ├── frame = data["Sheet1"]
  │     ├── iter_rows → headers + data ├── mask = frame["id"] > 10
  │     ├── Filter: _matches_where()   ├── result = frame[mask]["name"]
  │     └── Project: select columns    └── itertuples → List[Tuple]
  │     │
  │     ▼
  │   ExecutionResult(action="SELECT", rows=[("Alice",), ("Bob",)],
  │                   description=[("name", None, ...)], rowcount=2)
  │
  ▼
ExcelCursor
  ├── self._results = result.rows
  ├── self.description = result.description
  ├── self.rowcount = result.rowcount
  └── [No auto-save: SELECT is read-only]
```

### 4.2 Transaction Flow

```
conn = ExcelConnection("data.xlsx", autocommit=False)

State: initial_snapshot = engine.snapshot()
  │
  ▼
cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (1, 'Alice')")
  ├── In-memory data modified (row appended)
  └── No disk write (autocommit=False)
  │
  ▼
cursor.execute("UPDATE Sheet1 SET name = 'Bob' WHERE id = 1")
  ├── In-memory data modified (cell updated)
  └── No disk write
  │
  ├── Option A: conn.commit()
  │   ├── engine.save()           → Write to disk (atomic)
  │   └── new_snapshot = engine.snapshot()  → New rollback point
  │
  └── Option B: conn.rollback()
      └── engine.restore(initial_snapshot)  → Discard all changes
```

### 4.3 File I/O Flow

#### OpenpyxlEngine Save (Atomic):
```
engine.save()
  │
  ├── 1. directory = os.path.dirname(file_path)
  ├── 2. temp = tempfile.NamedTemporaryFile(dir=directory, suffix=".xlsx")
  ├── 3. workbook.save(temp.name)
  ├── 4. os.replace(temp.name, file_path)    ← ATOMIC (same filesystem)
  └── 5. finally: cleanup temp if replace failed
```

#### PandasEngine Save (Atomic):
```
engine.save()
  │
  ├── 1. directory = os.path.dirname(file_path)
  ├── 2. temp = tempfile.NamedTemporaryFile(dir=directory, suffix=".xlsx")
  ├── 3. pd.ExcelWriter(temp.name, engine="openpyxl")
  │     └── for sheet_name, frame in data.items():
  │           frame.to_excel(writer, sheet_name=sheet_name, index=False)
  ├── 4. os.replace(temp.name, file_path)    ← ATOMIC
  └── 5. finally: cleanup temp if replace failed
```

### 4.4 Snapshot / Restore Flow

#### OpenpyxlEngine:
```
snapshot():
  BytesIO buffer ← workbook.save(buffer) ← full workbook serialization

restore(snapshot):
  snapshot.seek(0)
  workbook = load_workbook(snapshot)       ← deserialize from buffer
  data = {sheet: workbook[sheet] for sheet in workbook.sheetnames}
```

#### PandasEngine:
```
snapshot():
  {name: frame.copy(deep=True) for name, frame in data.items()}

restore(snapshot):
  data = {name: frame.copy(deep=True) for name, frame in snapshot.items()}
```

---

## 5. Error Handling Architecture

### 5.1 Exception Mapping

| Source | Exception Type | When |
|--------|---------------|------|
| `check_closed` decorator | `InterfaceError` | Operations on closed connection/cursor |
| `ExcelConnection.__init__` | `InterfaceError` | Unsupported engine name |
| `ExcelConnection.rollback` | `NotSupportedError` | Rollback with autocommit=True |
| `ExcelConnection.workbook` | `NotSupportedError` | PandasEngine has no workbook property |
| `ExcelCursor.execute` | `ProgrammingError` | `ValueError` from parser or executor |
| `ExcelCursor.execute` | `NotSupportedError` | `NotImplementedError` from executor |
| Parser | `ValueError` | Invalid SQL syntax, missing params, extra params |
| Executor | `ValueError` | Sheet not found, unknown column, data mismatch |
| Executor | `NotImplementedError` | Unsupported WHERE operator |
| Engine save | `ValueError` | Workbook not loaded |

### 5.2 Exception Flow

```
User calls cursor.execute("BAD SQL")
  │
  ▼
ExcelCursor.execute()
  │
  ├── engine.execute_with_params() raises ValueError("Invalid SQL")
  │
  ▼
except ValueError as exc:
    raise ProgrammingError(str(exc)) from exc
  │
  ▼
User receives: ProgrammingError("Invalid SQL query format: BAD SQL")
```

---

## 6. Integration Architecture

### 6.1 sqlalchemy-excel Integration

excel-dbapi serves as the foundational I/O layer for sqlalchemy-excel:

```
┌──────────────────────────────────────────────────────────────┐
│  sqlalchemy-excel                                             │
│                                                               │
│  ExcelWorkbookSession                                         │
│  ┌───────────────────────────────────────────────────┐        │
│  │  self._conn = excel_dbapi.connect(                │        │
│  │      path, engine="openpyxl", create=True         │        │
│  │  )                                                │        │
│  │                                                   │        │
│  │  SQL Channel:     self._conn.cursor().execute()   │        │
│  │  Workbook Channel: self._conn.workbook            │        │
│  └───────────────────────────────────────────────────┘        │
│                                                               │
│  ExcelDbapiReader                                             │
│  ┌───────────────────────────────────────────────────┐        │
│  │  cursor = conn.cursor()                           │        │
│  │  cursor.execute("SELECT * FROM SheetName")        │        │
│  │  rows = cursor.fetchall()                         │        │
│  │  headers = [desc[0] for desc in cursor.description]│       │
│  └───────────────────────────────────────────────────┘        │
└──────────────────────────────────────────────────────────────┘
```

### 6.2 Contract Guarantees

excel-dbapi guarantees these interfaces for downstream consumers:

1. `connect()` returns an `ExcelConnection` with standard DB-API methods
2. `conn.workbook` returns openpyxl `Workbook` (openpyxl engine only)
3. `cursor.description` returns PEP 249 7-tuple column metadata
4. `cursor.fetchall()` returns `List[Tuple]`
5. `create=True` creates a valid empty workbook
6. All exceptions follow PEP 249 hierarchy

---

## 7. Performance Considerations

### 7.1 Memory Model

| Engine | In-Memory Representation | Memory Characteristics |
|--------|-------------------------|----------------------|
| OpenpyxlEngine | openpyxl Worksheet objects | Cell-level access, lazy loading available with `read_only` mode (not used) |
| PandasEngine | pandas DataFrames | Full data copy in memory, higher memory but faster vectorized ops |

### 7.2 Snapshot Cost

| Engine | Snapshot Method | Cost |
|--------|----------------|------|
| OpenpyxlEngine | Serialize workbook to BytesIO | O(n) time + O(n) memory (full workbook copy) |
| PandasEngine | Deep copy all DataFrames | O(n) time + O(n) memory (full data copy) |

### 7.3 Write Performance

| Engine | Save Method | Characteristics |
|--------|-------------|-----------------|
| OpenpyxlEngine | Atomic via tempfile + os.replace | Preserves formatting, charts, formulas |
| PandasEngine | Rewrite via pd.ExcelWriter | **Drops** formatting, charts, formulas |

### 7.4 Known Bottlenecks

1. **Large file initial load**: Both engines load entire workbook into memory
2. **Snapshot frequency**: Each commit creates a new snapshot — expensive for frequent commits
3. **DELETE with openpyxl**: `ws.delete_rows()` is O(n) per row — slow for bulk deletes
4. **No streaming**: All data loaded at once, no iterator-based streaming

---

## 8. Security Considerations

1. **No SQL injection risk**: Parameters are bound via `_bind_params()`, not string interpolation
2. **File access**: Only local filesystem access (no network URLs)
3. **XML attacks**: openpyxl handles XML parsing but does NOT use defusedxml by default — consumers should install defusedxml for untrusted files
4. **Formula injection**: `data_only=True` reads cached values, not formulas — mitigates formula injection in downstream display

---

## 9. Future Architecture (v2.0.x)

### 9.1 Planned Extensions

```
┌─────────────────────────────────────────────────┐
│  Layer 0: SQLAlchemy Integration (planned)       │
│  SQLAlchemy Dialect → uses DB-API interface      │
├─────────────────────────────────────────────────┤
│  Layer 1: DB-API 2.0 Interface (current)         │
├─────────────────────────────────────────────────┤
│  Layer 2: SQL Parsing (enhanced)                 │
│  + JOIN support, GROUP BY, HAVING, subqueries    │
├────────────────────┬────────────────┬───────────┤
│  OpenpyxlEngine    │ PandasEngine   │ Polars    │
│  (current)         │ (current)      │ (planned) │
└────────────────────┴────────────────┴───────────┘
```

### 9.2 SQLAlchemy Dialect Architecture (Planned)

```python
# Planned usage:
from sqlalchemy import create_engine
engine = create_engine("excel:///path/to/data.xlsx")

# The dialect would:
# 1. Implement SQLAlchemy's Dialect interface
# 2. Use excel-dbapi's connect() as the DBAPI
# 3. Map SQLAlchemy's SQL AST to excel-dbapi's SQL subset
```
