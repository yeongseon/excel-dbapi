from __future__ import annotations

from typing import Any

from hypothesis import given, settings, strategies as st

from excel_dbapi.parser import _tokenize, parse_sql


IDENTIFIER = st.from_regex(r"[A-Za-z_][A-Za-z0-9_]{0,8}", fullmatch=True)
NON_RESERVED_IDENTIFIER = IDENTIFIER.filter(
    lambda value: value.upper()
    not in {
        "SELECT",
        "FROM",
        "WHERE",
        "AND",
        "OR",
        "ORDER",
        "BY",
        "LIMIT",
        "OFFSET",
        "NULL",
        "IN",
        "NOT",
        "BETWEEN",
        "IS",
    }
)
INT_LITERAL = st.integers(min_value=-100000, max_value=100000)


def _literal_sql(value: int) -> str:
    return str(value)


def _condition_to_sql(condition: dict[str, Any]) -> str:
    operator = str(condition["operator"]).upper()
    column = str(condition["column"])
    if operator in {"IS", "IS NOT"}:
        return f"{column} {operator} NULL"
    if operator in {"IN", "NOT IN"}:
        values = ", ".join(_literal_sql(int(item)) for item in condition["value"])
        return f"{column} {operator} ({values})"
    if operator in {"BETWEEN", "NOT BETWEEN"}:
        low, high = condition["value"]
        return f"{column} {operator} {_literal_sql(int(low))} AND {_literal_sql(int(high))}"
    return f"{column} {operator} {_literal_sql(int(condition['value']))}"


def _where_to_sql(where: dict[str, Any]) -> str:
    conditions = where["conditions"]
    conjunctions = where["conjunctions"]
    parts: list[str] = []
    for index, condition in enumerate(conditions):
        if index:
            parts.append(str(conjunctions[index - 1]))
        parts.append(_condition_to_sql(condition))
    return " ".join(parts)


def _select_ast_to_sql(ast: dict[str, Any]) -> str:
    columns = ast["columns"]
    rendered_columns = ", ".join(str(column) for column in columns)
    sql = f"SELECT {rendered_columns} FROM {ast['table']}"

    where = ast.get("where")
    if isinstance(where, dict):
        sql += f" WHERE {_where_to_sql(where)}"

    order_by = ast.get("order_by")
    if isinstance(order_by, list) and order_by:
        order_item = order_by[0]
        sql += f" ORDER BY {order_item['column']} {order_item['direction']}"

    limit = ast.get("limit")
    if isinstance(limit, int):
        sql += f" LIMIT {limit}"

    offset = ast.get("offset")
    if isinstance(offset, int):
        sql += f" OFFSET {offset}"

    return sql


def _strip_params(ast: dict[str, Any]) -> dict[str, Any]:
    copy = dict(ast)
    copy.pop("params", None)
    return copy


@st.composite
def simple_select_sql(draw: st.DrawFn) -> str:
    table = draw(IDENTIFIER)
    columns = draw(st.lists(IDENTIFIER, min_size=1, max_size=4, unique=True))

    sql = f"SELECT {', '.join(columns)} FROM {table}"

    if draw(st.booleans()):
        column = draw(st.sampled_from(columns))
        operator = draw(st.sampled_from(["=", "!=", "<", "<=", ">", ">="]))
        value = draw(INT_LITERAL)
        sql += f" WHERE {column} {operator} {_literal_sql(value)}"

    if draw(st.booleans()):
        order_column = draw(st.sampled_from(columns))
        direction = draw(st.sampled_from(["ASC", "DESC"]))
        sql += f" ORDER BY {order_column} {direction}"

    if draw(st.booleans()):
        sql += f" LIMIT {draw(st.integers(min_value=0, max_value=1000))}"

    if draw(st.booleans()):
        sql += f" OFFSET {draw(st.integers(min_value=0, max_value=1000))}"

    return sql


@st.composite
def where_condition_sql(draw: st.DrawFn) -> str:
    column = draw(NON_RESERVED_IDENTIFIER)
    simple_op = draw(st.sampled_from(["=", "!=", "<>", "<", "<=", ">", ">="]))
    simple_value = draw(INT_LITERAL)

    in_values = draw(st.lists(INT_LITERAL, min_size=1, max_size=5))
    in_sql = ", ".join(_literal_sql(value) for value in in_values)

    low = draw(INT_LITERAL)
    high = draw(INT_LITERAL)
    between_low, between_high = sorted((low, high))

    return draw(
        st.sampled_from(
            [
                f"{column} {simple_op} {_literal_sql(simple_value)}",
                f"{column} IN ({in_sql})",
                f"{column} NOT IN ({in_sql})",
                f"{column} BETWEEN {_literal_sql(between_low)} AND {_literal_sql(between_high)}",
                f"{column} IS NULL",
                f"{column} IS NOT NULL",
            ]
        )
    )


@settings(max_examples=120)
@given(simple_select_sql())
def test_valid_generated_select_parses_without_crashing(sql: str) -> None:
    parsed = parse_sql(sql)
    assert parsed["action"] == "SELECT"


@settings(max_examples=100)
@given(simple_select_sql())
def test_select_ast_round_trips(sql: str) -> None:
    first = parse_sql(sql)
    regenerated_sql = _select_ast_to_sql(first)
    second = parse_sql(regenerated_sql)
    assert _strip_params(first) == _strip_params(second)


@settings(max_examples=80)
@given(
    st.text(
        alphabet=st.sampled_from(list(" abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-가나다漢字")),
        min_size=1,
        max_size=20,
    ).filter(lambda name: name.strip() != "")
)
def test_quoted_unicode_and_spaced_names_do_not_crash_parser(name: str) -> None:
    escaped = name.replace('"', '""')
    sql = f'SELECT "{escaped}" FROM "table {escaped}"'
    parsed = parse_sql(sql)
    assert parsed["action"] == "SELECT"


@settings(max_examples=120)
@given(where_condition_sql(), NON_RESERVED_IDENTIFIER)
def test_valid_where_conditions_do_not_crash_parser(condition_sql: str, table: str) -> None:
    sql = f"SELECT * FROM {table} WHERE {condition_sql}"
    parsed = parse_sql(sql)
    assert parsed["action"] == "SELECT"


@settings(max_examples=120)
@given(
    st.lists(
        st.one_of(
            st.sampled_from(
                [
                    "SELECT",
                    "FROM",
                    "WHERE",
                    "AND",
                    "OR",
                    "ORDER",
                    "BY",
                    "LIMIT",
                    "OFFSET",
                    "(",
                    ")",
                    ",",
                    "=",
                    "!=",
                    "<=",
                    ">=",
                    "<",
                    ">",
                    "NULL",
                    "IN",
                    "NOT",
                    "BETWEEN",
                ]
            ),
            IDENTIFIER,
            INT_LITERAL.map(str),
            st.text(alphabet=st.characters(blacklist_categories=("Cs",)), min_size=0, max_size=8).map(
                lambda value: "'" + value.replace("'", "''") + "'"
            ),
        ),
        min_size=1,
        max_size=25,
    )
)
def test_tokenizer_handles_valid_sql_token_streams(tokens: list[str]) -> None:
    sql = " ".join(tokens)
    tokenized = _tokenize(sql)
    assert isinstance(tokenized, list)
