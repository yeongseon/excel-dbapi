class InterfaceError(Exception):
    """Raised when a database interface operation fails."""


class NotSupportedError(Exception):
    """Raised when a requested operation is not supported."""
