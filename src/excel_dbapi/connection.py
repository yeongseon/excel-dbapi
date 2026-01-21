from typing import Optional, Type

from .cursor import ExcelCursor
from .engine.base import BaseEngine
from .engine.openpyxl_engine import OpenpyxlEngine
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

    def __init__(self, file_path: str):
        """
        Initialize the connection with the Excel file.
        """
        self.file_path: str = file_path
        self.closed: bool = False

        self.engine: BaseEngine = OpenpyxlEngine(file_path)
        self.data = self.engine.load()

    @check_closed
    def cursor(self) -> ExcelCursor:
        return ExcelCursor(self)

    @check_closed
    def commit(self) -> None:
        raise NotSupportedError("Transactions are not supported")

    @check_closed
    def rollback(self) -> None:
        raise NotSupportedError("Transactions are not supported")

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
