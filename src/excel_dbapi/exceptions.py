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
]
