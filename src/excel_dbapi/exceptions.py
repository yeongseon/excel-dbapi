class Error(Exception):
    """Base class for all DB-API errors."""


class Warning(Exception):
    """Raised for important warnings like data truncations."""


class InterfaceError(Error):
    """Raised when a database interface operation fails."""


class DatabaseError(Error):
    """Raised for errors related to the database."""


class DataError(DatabaseError):
    """Raised for data processing errors like value out of range."""


class OperationalError(DatabaseError):
    """Raised for operational errors like connection issues."""


class IntegrityError(DatabaseError):
    """Raised when relational integrity is affected."""


class InternalError(DatabaseError):
    """Raised when the database encounters an internal error."""


class ProgrammingError(DatabaseError):
    """Raised for programming errors like bad SQL syntax."""


class NotSupportedError(DatabaseError):
    """Raised when a requested operation is not supported."""


class SqlParseError(ProgrammingError):
    """SQL syntax error detected during parsing."""


class SqlSemanticError(ProgrammingError):
    """Valid SQL syntax but invalid semantics (e.g. unknown column, type mismatch)."""


class BackendOperationError(OperationalError):
    """Backend I/O or workbook operation failure."""


class CapabilityError(NotSupportedError):
    """Requested operation not supported by the current backend."""


__all__ = [
    "Error",
    "Warning",
    "InterfaceError",
    "DatabaseError",
    "DataError",
    "OperationalError",
    "IntegrityError",
    "InternalError",
    "ProgrammingError",
    "NotSupportedError",
    "SqlParseError",
    "SqlSemanticError",
    "BackendOperationError",
    "CapabilityError",
    "map_exception",
]


def map_exception(exc: Exception) -> "DatabaseError":
    """Map a non-DB-API exception to the most appropriate DB-API type.

    Used at the DB-API boundary to translate internal Python exceptions
    into PEP 249 exception types while preserving the original cause.
    """
    if isinstance(exc, Error):
        # Already a DB-API exception — should not be mapped.
        return exc  # type: ignore[return-value]
    if isinstance(exc, (ValueError, KeyError, TypeError, IndexError)):
        return ProgrammingError(str(exc))
    if isinstance(exc, NotImplementedError):
        return NotSupportedError(str(exc))
    if isinstance(exc, OSError):
        return OperationalError(str(exc))
    return DatabaseError(str(exc))
