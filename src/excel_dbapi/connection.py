import os
from collections.abc import Callable, Iterable, Sequence
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
    BackendOperationError,
    Error,
    InterfaceError,
    NotSupportedError,
    OperationalError,
    map_exception,
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
    if "://" in file_path and dsn_engine is None:
        scheme = file_path.split("://", 1)[0]
        raise OperationalError(
            f"Unsupported DSN scheme {scheme!r}. "
            f"Supported schemes: msgraph, sharepoint, onedrive"
        )
    if engine is None:
        engine = dsn_engine or "openpyxl"
    else:
        engine = engine.lower()
        if dsn_engine and engine != dsn_engine:
            raise BackendOperationError(
                f"Engine mismatch: DSN implies {dsn_engine!r}, got {engine!r}"
            )
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
        backup: bool = False,
        backup_dir: str | None = None,
        data_only: bool = True,
        sanitize_formulas: bool = True,
        credential: Credential = None,
        warn_rows: int | None = None,
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
            backup: If True, create a timestamped backup before the first
                mutating operation for local workbook files.
            backup_dir: Optional backup directory path. If omitted, backups are
                written to ``.excel-dbapi-backups`` under the workbook directory.
            data_only: If True, read cached formula values instead of formulas.
            sanitize_formulas: If True (default), escape cell values that could
                be interpreted as formulas by spreadsheet applications.
                This defends against formula injection (OWASP CSV Injection).
            credential: Optional credential / token provider for cloud backends.
            **backend_options: Extra keyword arguments forwarded to the backend.
        """
        self._data_only = data_only
        self._data_only_warning_issued = False
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
                "The pandas engine rewrites workbooks on save. "
                "Formatting, charts, images, and formulas will be dropped. "
                "Use engine='openpyxl' if you need to preserve these.",
                UserWarning,
                stacklevel=2,
            )

        # ── File existence check (local files only) ─────────────
        is_dsn = resolve_engine_from_dsn(file_path) is not None
        if not is_dsn:
            default_backup_dir = Path(self.file_path).parent / ".excel-dbapi-backups"
            requested_backup_dir = (
                Path(backup_dir).expanduser() if backup_dir is not None else default_backup_dir
            )
            if not requested_backup_dir.is_absolute():
                requested_backup_dir = (
                    Path(self.file_path).parent / requested_backup_dir
                )
            self._backup_dir: str | None = str(requested_backup_dir.resolve())
        else:
            self._backup_dir = None
        self._backup_enabled: bool = backup and not is_dsn
        self._backup_created: bool = False

        if not is_dsn and not create and not os.path.exists(self.file_path):
            raise OperationalError(f"Excel file not found: {self.file_path!r}")

        # ── Build backend options ────────────────────────────────
        opts: dict[str, Any] = {**backend_options}
        opts["warn_rows"] = warn_rows
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
            if isinstance(exc, Error):
                raise
            raise OperationalError(str(exc)) from exc

        # Guard: non-transactional backends reject autocommit=False
        if not autocommit and not self.engine.supports_transactions:
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
        self._snapshot: Any | None = None
        if not self.autocommit:
            self._snapshot = self.engine.snapshot()

    @check_closed
    def cursor(self) -> Any:
        from .cursor import ExcelCursor

        return ExcelCursor(self)

    @check_closed
    def commit(self) -> None:
        try:
            self._warn_data_only_if_needed()
            self.engine.save()
            self._snapshot = self.engine.snapshot()
        except Error:
            raise
        except Exception as exc:
            raise OperationalError(str(exc)) from exc

    @check_closed
    def rollback(self) -> None:
        try:
            if not self.engine.supports_transactions:
                raise NotSupportedError(
                    f"Backend '{self.engine.__class__.__name__}' does not support "
                    f"rollback (non-transactional backend)"
                )
            if self.autocommit:
                raise NotSupportedError(
                    "Rollback is disabled when autocommit is enabled"
                )
            if self._snapshot is None:
                raise NotSupportedError(
                    "No snapshot available for rollback"
                )
            self.engine.restore(self._snapshot)
        except Error:
            raise
        except Exception as exc:
            raise OperationalError(str(exc)) from exc

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
        except Error:
            raise
        except Exception as exc:
            raise map_exception(exc) from exc

    @check_closed
    def executemany(
        self, query: str, seq_of_params: Iterable[Sequence[Any]]
    ) -> ExecutionResult:
        """Execute *query* once for each parameter set in *seq_of_params*.

        Owns snapshot/restore orchestration so that transactional backends
        get atomic batch semantics.  Non-transactional backends warn on
        partial failure.
        """
        self._ensure_write_lock_for_query(query)
        supports_transactions = self.engine.supports_transactions
        snapshot = self.engine.snapshot() if supports_transactions else None
        backend_name = type(self.engine).__name__

        total_rowcount = 0
        last_rowid: int | None = None
        last_action: str | None = None

        for params in seq_of_params:
            try:
                result = self._executor.execute_with_params(query, tuple(params))
            except Error as exc:
                if supports_transactions:
                    self.engine.restore(snapshot)
                    raise
                raise type(exc)(
                    f"{exc}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            except Exception as exc:
                mapped = map_exception(exc)
                if supports_transactions:
                    self.engine.restore(snapshot)
                    raise mapped from exc
                raise type(mapped)(
                    f"{mapped}. Backend '{backend_name}' does not support transactional "
                    "executemany rollback; partial writes may have occurred."
                ) from exc
            total_rowcount += result.rowcount
            last_rowid = result.lastrowid
            last_action = result.action

        if self.autocommit and last_action is not None:
            self._finalize_autocommit(last_action)

        return ExecutionResult(
            action=last_action or "",
            rows=[],
            description=[],
            rowcount=total_rowcount,
            lastrowid=last_rowid,
        )

    def _ensure_write_lock_for_query(self, query: str) -> None:
        action = query.strip().split(None, 1)[0].upper() if query.strip() else ""
        if action in _MUTATING_ACTIONS:
            self.engine.ensure_write_lock()
            self._create_backup_if_needed()

    def _create_backup_if_needed(self) -> None:
        if not self._backup_enabled or self._backup_created:
            return
        if self._backup_dir is None:
            return
        source = Path(self.file_path)
        if not source.exists() or not source.is_file():
            return

        import shutil
        from datetime import datetime

        backup_dir = Path(self._backup_dir)
        backup_dir.mkdir(parents=True, exist_ok=True)
        stem = source.stem
        suffix = source.suffix
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        dest = backup_dir / f"{stem}.{timestamp}{suffix}"
        shutil.copy2(source, dest)
        self._backup_created = True

    def _finalize_autocommit(self, action: str) -> None:
        """Save and snapshot if autocommit is enabled and action is mutating.

        ``action`` is expected to be an uppercase SQL verb (e.g. ``"INSERT"``).
        The parser guarantees uppercase; callers must not pass lowercase.
        """
        if self.autocommit and action in _MUTATING_ACTIONS:
            self._warn_data_only_if_needed()
            self.engine.save()
            self._snapshot = self.engine.snapshot()
    def _warn_data_only_if_needed(self) -> None:
        if self._data_only and not self._data_only_warning_issued:
            # Compute stacklevel dynamically so the warning points at user code
            # regardless of whether the call comes through cursor or connection.
            import sys
            frame = sys._getframe(1)
            level = 2  # start at our caller
            pkg = __name__.rsplit(".", 1)[0]
            while frame.f_back is not None:
                module = frame.f_globals.get("__name__", "")
                if not module.startswith(pkg):
                    break
                frame = frame.f_back
                level += 1
            warnings.warn(
                "Workbook was opened with data_only=True (the default). "
                "Saving will replace formulas with their last cached values. "
                "Use connect(..., data_only=False) to preserve formulas.",
                UserWarning,
                stacklevel=level,
            )
            self._data_only_warning_issued = True

    def close(self) -> None:
        try:
            self.engine.close()
            self.closed = True
        except Error:
            raise
        except Exception as exc:
            raise OperationalError(str(exc)) from exc

    _CANONICAL_NAMES: dict[str, str] = {
        "OpenpyxlBackend": "openpyxl",
        "PandasBackend": "pandas",
        "GraphBackend": "graph",
    }

    @property
    def engine_name(self) -> str:
        cls_name = self.engine.__class__.__name__
        return self._CANONICAL_NAMES.get(cls_name, cls_name)

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
