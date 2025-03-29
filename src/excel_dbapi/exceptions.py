class Error(Exception):
    """Base class for all DBAPI exceptions."""

    pass


class InterfaceError(Error):
    """Raised for errors related to the database interface."""

    pass


class DatabaseError(Error):
    """Raised for errors related to the database itself."""

    pass


class TableError(Exception):
    """Raised when a table-related error occurs."""

    pass


class OperationalError(DatabaseError):
    """Raised for operational errors (e.g., connection issues)."""

    pass


class ProgrammingError(DatabaseError):
    """Raised for programming errors (e.g., invalid SQL)."""

    pass
