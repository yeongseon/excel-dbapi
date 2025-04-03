class Error(Exception):
    """Base class for all database-related exceptions."""

    pass


class DatabaseError(Error):
    """Exception for errors related to the database."""

    pass


class InterfaceError(Error):
    """Exception for errors related to the database interface."""

    pass


class DataError(DatabaseError):
    """Exception for errors due to problems with the processed data."""

    pass


class OperationalError(DatabaseError):
    """Exception for errors related to the database's operation."""

    pass


class IntegrityError(DatabaseError):
    """Exception for errors related to data integrity."""

    pass


class InternalError(DatabaseError):
    """Exception for internal database errors."""

    pass


class ProgrammingError(DatabaseError):
    """Exception for programming errors (SQL syntax, etc.)."""

    pass


class NotSupportedError(DatabaseError):
    """Exception for unsupported features."""

    pass
