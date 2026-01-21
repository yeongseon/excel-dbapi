from .connection import ExcelConnection

apilevel = "2.0"
threadsafety = 1
paramstyle = "qmark"


def connect(file_path: str) -> ExcelConnection:
    return ExcelConnection(file_path)


__all__ = [
    "ExcelConnection",
    "connect",
    "apilevel",
    "threadsafety",
    "paramstyle",
]
