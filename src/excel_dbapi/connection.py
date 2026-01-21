from typing import Any, Optional, Type

from .cursor import ExcelCursor
from .engine.base import BaseEngine
from .engine.openpyxl_engine import OpenpyxlEngine
from .engine.pandas_engine import PandasEngine
from .exceptions import InterfaceError, NotSupportedError


def check_closed(func):
    """Decorator to check if connection is closed before executing method."""
    def wrapper(self, *args, **kwargs):
        if self.closed:
            raise InterfaceError("Connection is already closed")
        return func(self, *args, **kwargs)
    return wrapper


class ExcelConnection:
    """
    ExcelConnection provides a PEP 249 compliant Connection interface
    for reading and querying Excel files using openpyxl.
    """

    def __init__(self, file_path: str, engine: str = "openpyxl", autocommit: bool = True):
        """
        Initialize the connection with the Excel file.
        """
        self.file_path: str = file_path
        self.closed: bool = False
        self.autocommit: bool = autocommit

        if engine == "openpyxl":
            self.engine = OpenpyxlEngine(file_path)
        elif engine == "pandas":
            self.engine = PandasEngine(file_path)
        else:
            raise InterfaceError(f"Unsupported engine: {engine}")

        self._snapshot: Any = self.engine.snapshot()

    @check_closed
    def cursor(self) -> ExcelCursor:
        return ExcelCursor(self)

    @check_closed
    def commit(self) -> None:
        self.engine.save()
        self._snapshot = self.engine.snapshot()

    @check_closed
    def rollback(self) -> None:
        if self.autocommit:
            raise NotSupportedError("Rollback is disabled when autocommit is enabled")
        self.engine.restore(self._snapshot)

    def close(self) -> None:
        self.closed = True

    @property
    def engine_name(self) -> str:
        return self.engine.__class__.__name__

    def __str__(self) -> str:
        return f"<ExcelConnection file='{self.file_path}' engine='{self.engine_name}' closed={self.closed}>"

    def __repr__(self) -> str:
        return self.__str__()

    def __enter__(self) -> "ExcelConnection":
        return self

    def __exit__(
        self,
        exc_type: Optional[Type[BaseException]],
        exc_val: Optional[BaseException],
        exc_tb,
    ) -> None:
        self.close()
