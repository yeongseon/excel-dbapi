import datetime
from typing import Any

from .connection import Credential, ExcelConnection
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
__version__ = "0.4.1"


class DBAPITypeObject:
    def __init__(self, *values: object) -> None:
        self.values = values

    def __eq__(self, other: object) -> bool:
        return other in self.values


def Date(year: int, month: int, day: int) -> datetime.date:
    return datetime.date(year, month, day)


def Time(hour: int, minute: int, second: int) -> datetime.time:
    return datetime.time(hour, minute, second)


def Timestamp(
    year: int,
    month: int,
    day: int,
    hour: int,
    minute: int,
    second: int,
) -> datetime.datetime:
    return datetime.datetime(year, month, day, hour, minute, second)


def DateFromTicks(ticks: float) -> datetime.date:
    return datetime.date.fromtimestamp(ticks)


def TimeFromTicks(ticks: float) -> datetime.time:
    return datetime.datetime.fromtimestamp(ticks).time()


def TimestampFromTicks(ticks: float) -> datetime.datetime:
    return datetime.datetime.fromtimestamp(ticks)


def Binary(string: Any) -> bytes:
    if isinstance(string, str):
        return string.encode()
    return bytes(string)


STRING = DBAPITypeObject(str)
BINARY = DBAPITypeObject(bytes, bytearray, memoryview)
NUMBER = DBAPITypeObject(int, float, bool)
DATETIME = DBAPITypeObject(datetime.date, datetime.time, datetime.datetime)
ROWID = DBAPITypeObject(int)


def connect(
    file_path: str,
    engine: str | None = None,
    autocommit: bool = True,
    create: bool = False,
    backup: bool = False,
    backup_dir: str | None = None,
    data_only: bool = True,
    sanitize_formulas: bool = True,
    credential: Credential = None,
    warn_rows: int | None = None,
    **backend_options: Any,
) -> ExcelConnection:
    return ExcelConnection(
        file_path,
        engine=engine,
        autocommit=autocommit,
        create=create,
        backup=backup,
        backup_dir=backup_dir,
        data_only=data_only,
        sanitize_formulas=sanitize_formulas,
        credential=credential,
        warn_rows=warn_rows,
        **backend_options,
    )


__all__ = [
    "ExcelConnection",
    "connect",
    "apilevel",
    "threadsafety",
    "paramstyle",
    "Date",
    "Time",
    "Timestamp",
    "DateFromTicks",
    "TimeFromTicks",
    "TimestampFromTicks",
    "Binary",
    "STRING",
    "BINARY",
    "NUMBER",
    "DATETIME",
    "ROWID",
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
