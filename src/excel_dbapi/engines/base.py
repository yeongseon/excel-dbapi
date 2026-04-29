from abc import ABC, abstractmethod
from dataclasses import dataclass
import errno
import os
from typing import Any
import warnings

from ..exceptions import BackendOperationError


@dataclass
class TableData:
    headers: list[str]
    rows: list[list[Any]]


class WorkbookBackend(ABC):
    file_path: str
    create: bool
    sanitize_formulas: bool

    @property
    @abstractmethod
    def readonly(self) -> bool:
        """Whether the backend is read-only (no write operations allowed)."""
        ...

    @property
    @abstractmethod
    def supports_transactions(self) -> bool:
        """Whether the backend supports commit/rollback transactions."""
        ...

    def __init__(
        self,
        file_path: str,
        *,
        data_only: bool = True,
        create: bool = False,
        sanitize_formulas: bool = True,
        **options: Any,
    ) -> None:
        self.file_path = file_path
        self.create = create
        self.sanitize_formulas = sanitize_formulas
        self.warn_rows: int | None = self._normalize_warn_rows(options.get("warn_rows"))
        self.max_rows: int | None = self._normalize_max_rows(options.get("max_rows"))
        self.max_memory_mb: float | None = self._normalize_max_memory_mb(
            options.get("max_memory_mb")
        )
        _is_local_path = "://" not in file_path
        self._file_locking_enabled = bool(options.get("file_locking", _is_local_path))
        self._lock_fd: int | None = None
        self._lock_path = f"{self.file_path}.lock"
        self._warn_rows_emitted: set[str] = set()
        self._row_warning_emitted: set[tuple[str, int]] = set()
        self._memory_warning_emitted: set[tuple[str, int]] = set()

    @staticmethod
    def _normalize_warn_rows(value: Any) -> int | None:
        if value is None:
            return None
        if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
            raise BackendOperationError("warn_rows must be a positive integer")
        return int(value)

    @staticmethod
    def _normalize_max_rows(value: Any) -> int | None:
        if value is None:
            return None
        if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
            raise BackendOperationError("max_rows must be a positive integer")
        return int(value)

    @staticmethod
    def _normalize_max_memory_mb(value: Any) -> float | None:
        if value is None:
            return None
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or float(value) <= 0
        ):
            raise BackendOperationError("max_memory_mb must be a positive number")
        return float(value)

    def _acquire_lock(self) -> None:
        """Acquire an advisory PID-based file lock.

        The lock is file-based: a ``<workbook>.lock`` file is created
        containing the owning process's PID.  If a stale lock is detected
        (PID no longer running), it is automatically cleared.

        .. note::
           This is **advisory locking** — it relies solely on PID
           existence checks and does not verify hostname or process
           start time.  In environments with rapid PID reuse, a stale
           lock may incorrectly appear active.
        """
        from ..exceptions import OperationalError

        if not self._file_locking_enabled or self._lock_fd is not None:
            return

        for _ in range(2):
            try:
                lock_fd = os.open(
                    self._lock_path,
                    os.O_CREAT | os.O_EXCL | os.O_WRONLY,
                    0o600,
                )
                os.write(lock_fd, str(os.getpid()).encode("ascii", "strict"))
                self._lock_fd = lock_fd
                return
            except OSError as exc:
                if exc.errno != errno.EEXIST:
                    raise OperationalError(str(exc)) from exc
                if not self._clear_stale_lock():
                    raise OperationalError("File is locked by another process") from exc

    def _clear_stale_lock(self) -> bool:
        try:
            with open(
                self._lock_path, "r", encoding="ascii", errors="replace"
            ) as handle:
                raw_pid = handle.read().strip()
        except (OSError, UnicodeDecodeError):
            return False

        try:
            lock_pid = int(raw_pid)
        except ValueError:
            lock_pid = -1

        is_stale = lock_pid <= 0
        if lock_pid > 0:
            try:
                os.kill(lock_pid, 0)
            except OSError as exc:
                is_stale = exc.errno == errno.ESRCH
            else:
                is_stale = False

        if not is_stale:
            return False

        try:
            os.unlink(self._lock_path)
        except FileNotFoundError:
            return True
        except OSError:
            return False
        return True

    def _release_lock(self) -> None:
        if self._lock_fd is None:
            return

        lock_fd = self._lock_fd
        self._lock_fd = None
        os.close(lock_fd)
        try:
            os.unlink(self._lock_path)
        except FileNotFoundError:
            pass

    def ensure_write_lock(self) -> None:
        if not self._file_locking_enabled:
            return
        self._acquire_lock()

    @staticmethod
    def _user_stacklevel() -> int:
        """Compute stacklevel that exits the excel_dbapi package."""
        import sys
        frame = sys._getframe(1)
        level = 2
        pkg = "excel_dbapi"
        while frame.f_back is not None:
            module = frame.f_globals.get("__name__", "")
            if not module.startswith(pkg):
                break
            frame = frame.f_back
            level += 1
        return level
    def _check_row_limit(self, sheet_name: str, row_count: int) -> None:
        from ..exceptions import OperationalError

        if self.warn_rows is not None and row_count > self.warn_rows:
            if sheet_name not in self._warn_rows_emitted:
                warnings.warn(
                    (
                        f"Sheet '{sheet_name}' has {row_count} rows. "
                        "excel-dbapi is optimized for small to medium workbooks. "
                        "For large analytical workloads, consider SQLite, DuckDB, or PostgreSQL."
                    ),
                    UserWarning,
                    stacklevel=self._user_stacklevel(),
                )
                self._warn_rows_emitted.add(sheet_name)

        if self.max_rows is None:
            return

        warning_threshold = max(1, int(self.max_rows * 0.8))
        warning_key = (sheet_name, self.max_rows)
        if (
            warning_key not in self._row_warning_emitted
            and warning_threshold <= row_count <= self.max_rows
        ):
            warnings.warn(
                f"Sheet '{sheet_name}' has reached {row_count}/{self.max_rows} rows",
                stacklevel=self._user_stacklevel(),
            )
            self._row_warning_emitted.add(warning_key)
        if row_count > self.max_rows:
            raise OperationalError("Sheet exceeds max_rows limit")

    def _check_memory_limit(self, sheet_name: str, approx_bytes: int) -> None:
        from ..exceptions import OperationalError

        if self.max_memory_mb is None:
            return

        limit_bytes = int(self.max_memory_mb * 1024 * 1024)
        warning_threshold = max(1, int(limit_bytes * 0.8))
        warning_key = (sheet_name, limit_bytes)
        if (
            warning_key not in self._memory_warning_emitted
            and warning_threshold <= approx_bytes <= limit_bytes
        ):
            warnings.warn(
                (
                    f"Sheet '{sheet_name}' has reached approximately "
                    f"{approx_bytes / (1024 * 1024):.2f}/{self.max_memory_mb:.2f} MB"
                ),
                stacklevel=self._user_stacklevel(),
            )
            self._memory_warning_emitted.add(warning_key)
        if approx_bytes > limit_bytes:
            raise OperationalError("Sheet exceeds max_memory_mb limit")

    @abstractmethod
    def load(self) -> None:
        pass

    @abstractmethod
    def save(self) -> None:
        pass

    @abstractmethod
    def snapshot(self) -> Any:
        pass

    @abstractmethod
    def restore(self, snapshot: Any) -> None:
        pass

    @abstractmethod
    def list_sheets(self) -> list[str]:
        pass

    @abstractmethod
    def read_sheet(self, sheet_name: str) -> TableData:
        pass

    @abstractmethod
    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        pass

    @abstractmethod
    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        pass

    @abstractmethod
    def create_sheet(self, name: str, headers: list[str]) -> None:
        pass

    @abstractmethod
    def drop_sheet(self, name: str) -> None:
        pass

    def close(self) -> None:
        self._release_lock()

    def get_workbook(self) -> Any:
        from ..exceptions import NotSupportedError

        raise NotSupportedError(
            f"Backend '{type(self).__name__}' does not expose a workbook object"
        )


def _normalize_headers(raw: list[Any]) -> list[str]:
    """Validate and normalise raw header values to a list of strings.

    Raises
    ------
    DataError
        If any header is empty/None or if there are duplicate headers.
    """
    from ..exceptions import DataError

    headers: list[str] = []
    for idx, value in enumerate(raw):
        if value is None or (isinstance(value, str) and value.strip() == ""):
            raise DataError(f"Empty or None header at column index {idx}")
        headers.append(str(value).strip())

    seen: set[str] = set()
    lower_map: dict[str, str] = {}
    for h in headers:
        key = h.casefold()
        if key in seen:
            raise DataError(
                f"Duplicate header: {h!r} (conflicts with {lower_map[key]!r})"
            )
        seen.add(key)
        lower_map[key] = h

    return headers
