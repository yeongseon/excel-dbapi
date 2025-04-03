from typing import Optional, Type

from .cursor import ExcelCursor
from .engine import BaseEngine, OpenpyxlEngine, PandasEngine
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
    for reading and querying Excel files.
    """

    def __init__(self, file_path: str, engine: str = "openpyxl"):
        """
        Initialize the connection with the Excel file and selected engine.

        Args:
            file_path (str): Path to the Excel file.
            engine (str): Engine type ('pandas' or 'openpyxl').
        """
        self.file_path: str = file_path
        self.closed: bool = False

        if engine == "pandas":
            self.engine: BaseEngine = PandasEngine(file_path)
        elif engine == "openpyxl":
            self.engine: BaseEngine = OpenpyxlEngine(file_path)
        else:
            raise ValueError(f"Unsupported engine: {engine}")

    @check_closed
    def cursor(self) -> ExcelCursor:
        """
        Return a new Cursor object using the connection.

        Returns:
            ExcelCursor: A new cursor object.
        """
        return ExcelCursor(self.engine)

    @check_closed
    def commit(self) -> None:
        """
        Commit any pending transaction (Not supported for Excel).

        Raises:
            NotSupportedError: Always raised because transactions are not supported.
        """
        raise NotSupportedError("Transactions are not supported")

    @check_closed
    def rollback(self) -> None:
        """
        Roll back to the start of any pending transaction (Not supported for Excel).

        Raises:
            NotSupportedError: Always raised because transactions are not supported.
        """
        raise NotSupportedError("Transactions are not supported")

    def close(self) -> None:
        """
        Close the connection.
        """
        self.closed = True

    @property
    def engine_name(self) -> str:
        """
        Return the name of the engine being used.
        """
        return self.engine.__class__.__name__

    def __str__(self) -> str:
        return f"<ExcelConnection file='{self.file_path}' engine='{self.engine_name}' closed={self.closed}>"

    def __repr__(self) -> str:
        return self.__str__()

    def __enter__(self) -> "ExcelConnection":
        """
        Enter the runtime context related to this object.
        """
        return self

    def __exit__(
        self,
        exc_type: Optional[Type[BaseException]],
        exc_val: Optional[BaseException],
        exc_tb,
    ) -> None:
        """
        Exit the runtime context and close the connection.
        """
        self.close()
