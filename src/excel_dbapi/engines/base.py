from abc import ABC, abstractmethod
from dataclasses import dataclass
import errno
import os
from typing import Any


@dataclass
class TableData:
    headers: list[str]
    rows: list[list[Any]]


class WorkbookBackend(ABC):
    file_path: str
    create: bool
    sanitize_formulas: bool
    readonly: bool = False
    supports_transactions: bool = True

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
        _is_local_path = not ("://" in file_path)
        self._file_locking_enabled = bool(options.get("file_locking", _is_local_path))
        self._lock_fd: int | None = None
        self._lock_path = f"{self.file_path}.lock"

    def _acquire_lock(self) -> None:
        from ..exceptions import OperationalError

        if not self._file_locking_enabled or self._lock_fd is not None:
            return

        try:
            lock_fd = os.open(self._lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            os.write(lock_fd, str(os.getpid()).encode("ascii", "strict"))
            self._lock_fd = lock_fd
        except OSError as exc:
            if exc.errno == errno.EEXIST:
                raise OperationalError("File is locked by another process") from exc
            raise OperationalError(str(exc)) from exc

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
            raise DataError(
                f"Empty or None header at column index {idx}"
            )
        headers.append(str(value).strip())

    seen: set[str] = set()
    lower_map: dict[str, str] = {}
    for h in headers:
        key = h.lower()
        if key in seen:
            raise DataError(
                f"Duplicate header: {h!r} (conflicts with {lower_map[key]!r})"
            )
        seen.add(key)
        lower_map[key] = h

    return headers
