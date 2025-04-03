from .connection import ExcelConnection

def connect(filename, **kwargs):
    """
    DBAPI-compliant connect function for Excel files.
    """
    return ExcelConnection(filename, **kwargs).connect()

__all__ = ["connect"]