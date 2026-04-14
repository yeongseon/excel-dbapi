import importlib
import os
from collections.abc import Callable, Sequence
from functools import wraps
from pathlib import Path
from types import TracebackType
from typing import (
    Any,
    Concatenate,
    Optional,
    ParamSpec,
    Protocol,
    Type,
    TypeVar,
    cast,
    runtime_checkable,
)
import warnings

from .engines.base import WorkbookBackend
from .engines.registry import get_engine, resolve_engine_from_dsn
from .executor import SharedExecutor
from .engines.result import ExecutionResult
from .exceptions import (
    DatabaseError,
    InterfaceError,
    NotSupportedError,
    OperationalError,
    ProgrammingError,
)

try:
    from openpyxl.utils.exceptions import InvalidFileException as _InvalidFileException

    InvalidFileException: type[Exception] | None = _InvalidFileException
except ImportError:
    InvalidFileException = None


@runtime_checkable
class _TokenProvider(Protocol):
    """Object that supplies a bearer token string."""

    def get_token(self, *args: Any) -> Any: ...


#: Credential accepted by cloud backends.  Concrete forms:
#: ``str`` (static token), ``TokenProvider`` protocol,
#: azure-identity credential (``get_token(scope)``), or zero-arg callable.
Credential = str | Callable[[], str] | _TokenProvider | None

_MUTATING_ACTIONS = frozenset({"INSERT", "CREATE", "DROP", "UPDATE", "DELETE", "ALTER"})

P = ParamSpec("P")
R = TypeVar("R")


def check_closed(
    func: Callable[Concatenate["ExcelConnection", P], R],
) -> Callable[Concatenate["ExcelConnection", P], R]:
    """Decorator to check if connection is closed before executing method."""

    @wraps(func)
    def wrapper(self: "ExcelConnection", *args: P.args, **kwargs: P.kwargs) -> R:
        if self.closed:
            raise InterfaceError("Connection is already closed")
        return func(self, *args, **kwargs)

    return cast(Callable[Concatenate["ExcelConnection", P], R], wrapper)


def _resolve_engine_and_location(file_path: str, engine: str | None) -> tuple[str, str]:
    """Determine engine name and normalised location from file_path/DSN."""
    dsn_engine = resolve_engine_from_dsn(file_path)
    if engine is None:
        engine = dsn_engine or "openpyxl"
    elif dsn_engine and engine != dsn_engine:
        raise ValueError(f"Engine mismatch: DSN implies {dsn_engine!r}, got {engine!r}")
    if dsn_engine:
        return engine, file_path  # Don't Path.resolve() URLs/DSNs
    return engine, str(Path(file_path).expanduser().resolve())


class ExcelConnection:
    """
    ExcelConnection provides a PEP 249 compliant Connection interface
    for reading and querying Excel files using openpyxl.
    """

    def __init__(
        self,
        file_path: str,
        engine: str | None = None,
        autocommit: bool = True,
        create: bool = False,
        data_only: bool = True,
        sanitize_formulas: bool = True,
        credential: Credential = None,
        **backend_options: Any,
    ):
        """
        Initialize the connection with the Excel file.

        Args:
            file_path: Path to the Excel (.xlsx) file or a DSN
                (e.g. ``msgraph://drives/{id}/items/{id}``).
            engine: Engine backend name ("openpyxl", "pandas", "graph",
                or None for auto-detection from DSN).
            autocommit: If True, auto-save after write operations.
            create: If True, create the file if it does not exist.
            data_only: If True, read cached formula values instead of formulas.
            sanitize_formulas: If True (default), escape cell values that could
                be interpreted as formulas by spreadsheet applications.
                This defends against formula injection (OWASP CSV Injection).
            credential: Optional credential / token provider for cloud backends.
            **backend_options: Extra keyword arguments forwarded to the backend.
        """
        # ── Resolve engine + location ──────────────────────────────
        try:
            engine_name, location = _resolve_engine_and_location(file_path, engine)
        except ValueError as exc:
            raise OperationalError(str(exc)) from exc

        self.file_path: str = location
        self.closed: bool = False
        self.autocommit: bool = autocommit

        try:
            engine_cls = get_engine(engine_name)
        except ValueError as exc:
            raise OperationalError(str(exc)) from exc

        if engine_name == "pandas":
            warnings.warn(
                "The pandas engine will become an optional dependency in v2.0. "
                "Install with: pip install excel-dbapi[pandas]",
                DeprecationWarning,
                stacklevel=2,
            )

        # ── File existence check (local files only) ─────────────
        is_dsn = resolve_engine_from_dsn(file_path) is not None
        if not is_dsn and not create and not os.path.exists(self.file_path):
            raise OperationalError(f"Excel file not found: {self.file_path!r}")

        # ── Build backend options ────────────────────────────────
        opts: dict[str, Any] = {**backend_options}
        if credential is not None:
            opts["credential"] = credential

        # ── Instantiate backend with exception translation ────────────────
        try:
            self.engine: WorkbookBackend = engine_cls(
                self.file_path,
                data_only=data_only,
                create=create,
                sanitize_formulas=sanitize_formulas,
                **opts,
            )
        except (FileNotFoundError, ValueError) as exc:
            # Translate backend construction errors to OperationalError
            raise OperationalError(str(exc)) from exc
        except Exception as exc:
            # Catch and translate openpyxl's InvalidFileException,
            # zipfile.BadZipFile, and other file-format errors
            exc_class_name = type(exc).__name__
            if exc_class_name in (
                "InvalidFileException",
                "BadZipFile",
                "BadZipfile",
            ) or (
                InvalidFileException is not None
                and isinstance(exc, InvalidFileException)
            ):
                raise OperationalError(str(exc)) from exc
            # Re-raise other unexpected exceptions
            raise

        # Guard: non-transactional backends reject autocommit=False
        if not autocommit and not getattr(self.engine, "supports_transactions", True):
            self.engine.close()
            raise NotSupportedError(
                f"Backend '{self.engine.__class__.__name__}' does not support "
                f"transactions (autocommit=False)"
            )

        self._executor = SharedExecutor(
            self.engine,
            sanitize_formulas=sanitize_formulas,
            connection=self,
        )
        if not self.autocommit:
            self.engine.ensure_write_lock()
        self._snapshot: Any = self.engine.snapshot()

    @check_closed
    def cursor(self) -> Any:
        cursor_module = importlib.import_module("excel_dbapi.cursor")
        ExcelCursor = cursor_module.ExcelCursor
        return ExcelCursor(self)

    @check_closed
    def commit(self) -> None:
        self.engine.save()
        self._snapshot = self.engine.snapshot()

    @check_closed
    def rollback(self) -> None:
        if not getattr(self.engine, "supports_transactions", True):
            raise NotSupportedError(
                f"Backend '{self.engine.__class__.__name__}' does not support "
                f"rollback (non-transactional backend)"
            )
        if self.autocommit:
            raise NotSupportedError("Rollback is disabled when autocommit is enabled")
        self.engine.restore(self._snapshot)

    @check_closed
    def execute(
        self, query: str, params: Sequence[Any] | None = None
    ) -> ExecutionResult:
        try:
            self._ensure_write_lock_for_query(query)
            normalized_params = tuple(params) if params is not None else None
            result = self._executor.execute_with_params(query, normalized_params)
            self._finalize_autocommit(result.action)
            return result
        except ValueError as exc:
            raise ProgrammingError(str(exc)) from exc
        except NotImplementedError as exc:
            raise NotSupportedError(str(exc)) from exc
        except (KeyError, TypeError, IndexError) as exc:
            raise ProgrammingError(str(exc)) from exc
        except OSError as exc:
            raise OperationalError(str(exc)) from exc
        except DatabaseError:
            raise
        except Exception as exc:
            raise DatabaseError(str(exc)) from exc

    def _ensure_write_lock_for_query(self, query: str) -> None:
        action = query.strip().split(None, 1)[0].upper() if query.strip() else ""
        if action in _MUTATING_ACTIONS:
            self.engine.ensure_write_lock()

    def _finalize_autocommit(self, action: str) -> None:
        """Save and snapshot if autocommit is enabled and action is mutating.

        ``action`` is expected to be an uppercase SQL verb (e.g. ``"INSERT"``).
        The parser guarantees uppercase; callers must not pass lowercase.
        """
        if self.autocommit and action in _MUTATING_ACTIONS:
            self.engine.save()
            self._snapshot = self.engine.snapshot()

    def close(self) -> None:
        self.engine.close()
        self.closed = True

    @property
    def engine_name(self) -> str:
        return self.engine.__class__.__name__

    @property
    def workbook(self) -> Any:
        return self.engine.get_workbook()

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
        exc_tb: Optional[TracebackType],
    ) -> None:
        self.close()
