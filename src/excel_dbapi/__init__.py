from pathlib import Path

from .connection import ExcelConnection

apilevel = "2.0"
threadsafety = 1
paramstyle = "qmark"
__version__ = (Path(__file__).resolve().parents[2] / "VERSION").read_text(encoding="utf-8").strip()


def connect(
    file_path: str,
    engine: str = "openpyxl",
    autocommit: bool = True,
    create: bool = False,
    data_only: bool = True,
) -> ExcelConnection:
    return ExcelConnection(
        file_path,
        engine=engine,
        autocommit=autocommit,
        create=create,
        data_only=data_only,
    )


__all__ = [
    "ExcelConnection",
    "connect",
    "apilevel",
    "threadsafety",
    "paramstyle",
    "__version__",
]
