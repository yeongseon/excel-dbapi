import pytest
from excel_dbapi.exceptions import InterfaceError, NotSupportedError


def test_interface_error():
    with pytest.raises(InterfaceError):
        raise InterfaceError("Connection closed")


def test_not_supported_error():
    with pytest.raises(NotSupportedError):
        raise NotSupportedError("Not supported operation")
