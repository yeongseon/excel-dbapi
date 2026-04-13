from __future__ import annotations

from hypothesis import given, settings, strategies as st

from excel_dbapi.parser import _tokenize, parse_sql


ParseError = ValueError


@settings(max_examples=300)
@given(st.text(max_size=500))
def test_tokenizer_and_parser_do_not_crash_on_random_strings(text: str) -> None:
    _tokenize(text)
    try:
        parse_sql(text)
    except ParseError:
        pass


def test_tokenizer_handles_empty_string() -> None:
    assert _tokenize("") == []


def test_parser_empty_string_raises_parse_error() -> None:
    try:
        parse_sql("")
    except ParseError:
        return
    assert False, "Expected parse_sql to raise ValueError for empty input"


def test_tokenizer_handles_very_long_string() -> None:
    text = "SELECT " + ("a" * 20000)
    tokens = _tokenize(text)
    assert len(tokens) >= 2


def test_tokenizer_and_parser_handle_null_bytes() -> None:
    text = "SELECT\x00*\x00FROM\x00Sheet1"
    _tokenize(text)
    try:
        parse_sql(text)
    except ParseError:
        pass
