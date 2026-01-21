from .connection import ExcelConnection

apilevel = "2.0"
threadsafety = 1
paramstyle = "qmark"


def connect(file_path: str, engine: str = "openpyxl", autocommit: bool = True) -> ExcelConnection:
    return ExcelConnection(file_path, engine=engine, autocommit=autocommit)


__all__ = [
    "ExcelConnection",
    "connect",
    "apilevel",
    "threadsafety",
    "paramstyle",
]
