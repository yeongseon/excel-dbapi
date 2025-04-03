import pandas as pd
import pytest

from excel_dbapi.engine.executor import execute_query


def test_execute_select_query():
    data = {
        "Sheet1": pd.DataFrame([{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}])
    }
    parsed = {"action": "SELECT", "table": "Sheet1", "where": "id == 1"}
    result = execute_query(parsed, data)
    assert result == [{"id": 1, "name": "Alice"}]


def test_execute_invalid_table():
    data = {}
    parsed = {"action": "SELECT", "table": "Unknown"}
    with pytest.raises(ValueError):
        execute_query(parsed, data)


def test_execute_unsupported_action():
    data = {"Sheet1": pd.DataFrame()}
    parsed = {"action": "UPDATE", "table": "Sheet1"}
    with pytest.raises(NotImplementedError):
        execute_query(parsed, data)
