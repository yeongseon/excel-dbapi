from typing import Any

from .connection import ExcelConnection
from .exceptions import (
    DatabaseError,
    DataError,
    Error,
    IntegrityError,
    InterfaceError,
    InternalError,
    NotSupportedError,
    OperationalError,
    ProgrammingError,
    Warning,
)
from .reflection import (
    METADATA_SHEET,
    get_columns,
    has_table,
    list_tables,
    read_table_metadata,
    remove_table_metadata,
    write_table_metadata,
)

apilevel = "2.0"
threadsafety = 1
paramstyle = "qmark"
__version__ = "0.1.1"


def connect(
    file_path: str,
    engine: str | None = "openpyxl",
    autocommit: bool = True,
    create: bool = False,
    data_only: bool = True,
    sanitize_formulas: bool = True,
    credential: Any = None,
    **backend_options: Any,
) -> ExcelConnection:
    return ExcelConnection(
        file_path,
        engine=engine,
        autocommit=autocommit,
        create=create,
        data_only=data_only,
        sanitize_formulas=sanitize_formulas,
        credential=credential,
        **backend_options,
    )


__all__ = [
    "ExcelConnection",
    "connect",
    "apilevel",
    "threadsafety",
    "paramstyle",
    "__version__",
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
    "list_tables",
    "has_table",
    "get_columns",
    "read_table_metadata",
    "write_table_metadata",
    "remove_table_metadata",
    "METADATA_SHEET",
]
