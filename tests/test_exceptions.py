import pytest

from excel_dbapi.exceptions import (DatabaseError, DataError, Error,
                                    IntegrityError, InterfaceError,
                                    InternalError, NotSupportedError,
                                    OperationalError, ProgrammingError)


def test_exception_hierarchy():
    assert issubclass(DatabaseError, Error)
    assert issubclass(InterfaceError, Error)
    assert issubclass(DataError, DatabaseError)
    assert issubclass(OperationalError, DatabaseError)
    assert issubclass(IntegrityError, DatabaseError)
    assert issubclass(InternalError, DatabaseError)
    assert issubclass(ProgrammingError, DatabaseError)
    assert issubclass(NotSupportedError, DatabaseError)
