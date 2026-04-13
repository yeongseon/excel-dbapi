from typing import Any

from excel_dbapi.executor import SharedExecutor
from excel_dbapi.engines.base import TableData, WorkbookBackend
from excel_dbapi.parser import parse_sql


class _StubBackend(WorkbookBackend):
    def __init__(self) -> None:
        super().__init__("stub.xlsx")

    def load(self) -> None:
        return None

    def save(self) -> None:
        return None

    def snapshot(self) -> object:
        return object()

    def restore(self, snapshot: object) -> None:
        return None

    def list_sheets(self) -> list[str]:
        return []

    def read_sheet(self, sheet_name: str) -> TableData:
        raise NotImplementedError

    def write_sheet(self, sheet_name: str, data: TableData) -> None:
        raise NotImplementedError

    def append_row(self, sheet_name: str, row: list[Any]) -> int:
        raise NotImplementedError

    def create_sheet(self, name: str, headers: list[str]) -> None:
        raise NotImplementedError

    def drop_sheet(self, name: str) -> None:
        raise NotImplementedError


def _executor() -> SharedExecutor:
    return SharedExecutor(_StubBackend())


def test_unicode_headers_and_data_in_row_evaluation() -> None:
    executor = _executor()
    row = {
        "日本語": "東京",
        "中文": "北京",
        "한국어": "서울",
        "العربية": "مرحبا",
        "emoji_header_🎉": "📊 metrics",
    }
    assert executor._evaluate_condition(
        row,
        {"column": "日本語", "operator": "=", "value": "東京"},
    )
    assert executor._evaluate_condition(
        row,
        {"column": "emoji_header_🎉", "operator": "LIKE", "value": "📊%"},
    )


def test_unicode_where_predicate_parsing() -> None:
    parsed = parse_sql("SELECT id FROM t WHERE mixed = 'report-北京'")
    assert parsed["where"]["conditions"][0] == {
        "column": "mixed",
        "operator": "=",
        "value": "report-北京",
    }


def test_unicode_order_by_and_group_by_parsing() -> None:
    parsed = parse_sql(
        "SELECT group_key, COUNT(*) FROM t GROUP BY group_key ORDER BY group_key"
    )
    assert parsed["group_by"] == ["group_key"]
    assert parsed["order_by"][0]["column"] == "group_key"


def test_unicode_like_ilike_matching() -> None:
    executor = _executor()
    row = {"mixed": "Report-東京"}
    assert executor._evaluate_condition(
        row,
        {"column": "mixed", "operator": "LIKE", "value": "%東京"},
    )
    assert executor._evaluate_condition(
        row,
        {"column": "mixed", "operator": "ILIKE", "value": "report-%"},
    )


def test_full_width_and_normalization_edge_cases() -> None:
    executor = _executor()

    assert executor._evaluate_condition(
        {"width": "ＡＢＣ"},
        {"column": "width", "operator": "LIKE", "value": "ＡＢ_"},
    )

    assert executor._evaluate_condition(
        {"normalized": "café"},
        {"column": "normalized", "operator": "=", "value": "café"},
    )
    assert not executor._evaluate_condition(
        {"normalized": "café"},
        {"column": "normalized", "operator": "=", "value": "cafe\u0301"},
    )
