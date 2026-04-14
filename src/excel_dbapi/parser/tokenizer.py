import re
from typing import Any, List

from ._constants import (
    _AGGREGATE_FUNCTIONS,
    _IDENTIFIER_PATTERN,
    _QUALIFIED_IDENTIFIER_PATTERN,
    _QuotedString,
)


def _split_csv(text: str) -> List[str]:
    items: List[str] = []
    current: List[str] = []
    in_single = False
    in_double = False
    paren_depth = 0
    case_depth = 0
    index = 0

    while index < len(text):
        char = text[index]

        if in_single:
            current.append(char)
            if char == "'":
                if index + 1 < len(text) and text[index + 1] == "'":
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_single = False
            index += 1
            continue

        if in_double:
            current.append(char)
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_double = False
            index += 1
            continue

        if char == "'":
            in_single = True
            current.append(char)
            index += 1
            continue

        if char == '"':
            in_double = True
            current.append(char)
            index += 1
            continue

        if char == "(":
            paren_depth += 1
            current.append(char)
            index += 1
            continue

        if char == ")":
            if paren_depth > 0:
                paren_depth -= 1
            current.append(char)
            index += 1
            continue

        if char.isalpha() or char == "_":
            start = index
            while index < len(text) and (text[index].isalnum() or text[index] == "_"):
                index += 1
            word = text[start:index]
            upper = word.upper()
            if upper == "CASE":
                case_depth += 1
            elif upper == "END" and case_depth > 0:
                case_depth -= 1
            current.append(word)
            continue

        if char == "," and paren_depth == 0 and case_depth == 0:
            items.append("".join(current).strip())
            current = []
            index += 1
            continue

        current.append(char)
        index += 1

    if current:
        items.append("".join(current).strip())
    return items


def _split_csv_preserve_empty(text: str) -> List[str]:
    items: List[str] = []
    current: List[str] = []
    in_single = False
    in_double = False
    paren_depth = 0
    case_depth = 0
    index = 0

    while index < len(text):
        char = text[index]

        if in_single:
            current.append(char)
            if char == "'":
                if index + 1 < len(text) and text[index + 1] == "'":
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_single = False
            index += 1
            continue

        if in_double:
            current.append(char)
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_double = False
            index += 1
            continue

        if char == "'":
            in_single = True
            current.append(char)
            index += 1
            continue

        if char == '"':
            in_double = True
            current.append(char)
            index += 1
            continue

        if char == "(":
            paren_depth += 1
            current.append(char)
            index += 1
            continue

        if char == ")":
            if paren_depth > 0:
                paren_depth -= 1
            current.append(char)
            index += 1
            continue

        if char.isalpha() or char == "_":
            start = index
            while index < len(text) and (text[index].isalnum() or text[index] == "_"):
                index += 1
            word = text[start:index]
            upper = word.upper()
            if upper == "CASE":
                case_depth += 1
            elif upper == "END" and case_depth > 0:
                case_depth -= 1
            current.append(word)
            continue

        if char == "," and paren_depth == 0 and case_depth == 0:
            items.append("".join(current).strip())
            current = []
            index += 1
            continue

        current.append(char)
        index += 1

    items.append("".join(current).strip())
    return items


def _tokenize(text: str) -> List[str]:
    tokens: List[str] = []
    current: List[str] = []
    in_single = False
    in_double = False
    index = 0

    while index < len(text):
        char = text[index]

        if in_single:
            current.append(char)
            if char == "'":
                if index + 1 < len(text) and text[index + 1] == "'":
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_single = False
            index += 1
            continue

        if in_double:
            current.append(char)
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_double = False
            index += 1
            continue

        if char.isspace():
            if current:
                tokens.append("".join(current))
                current = []
            index += 1
            continue

        if char == "'":
            current.append(char)
            in_single = True
            index += 1
            continue

        if char == '"':
            current.append(char)
            in_double = True
            index += 1
            continue

        if char in {"(", ")"}:
            if current:
                tokens.append("".join(current))
                current = []
            tokens.append(char)
            index += 1
            continue

        current.append(char)
        index += 1

    if current:
        tokens.append("".join(current))

    return tokens


def _count_unquoted_placeholders(sql: str) -> int:
    """Count ``?`` placeholders outside string literals in *sql*."""
    count = 0
    in_quote = False
    quote_char = ""
    i = 0
    length = len(sql)
    while i < length:
        ch = sql[i]
        if in_quote:
            if ch == quote_char:
                if i + 1 < length and sql[i + 1] == quote_char:
                    i += 2
                    continue
                in_quote = False
        else:
            if ch in ("'", '"'):
                in_quote = True
                quote_char = ch
            elif ch == "?":
                count += 1
        i += 1
    return count


def _parse_value(token: str) -> Any:
    token = token.strip()
    if token.upper() == "NULL":
        return None
    if token.startswith("'") and token.endswith("'") and len(token) >= 2:
        # Unescape doubled single quotes: 'it''s' -> it's
        return _QuotedString(token[1:-1].replace("''", "'"))
    if token.startswith('"') and token.endswith('"') and len(token) >= 2:
        # Unescape doubled double quotes: "say ""hello""" -> say "hello"
        return _QuotedString(token[1:-1].replace('""', '"'))
    try:
        return int(token)
    except ValueError:
        pass
    try:
        return float(token)
    except ValueError:
        return token


def _parse_table_identifier(token: str) -> str:
    identifier = token.strip()
    if not identifier:
        raise ValueError("Table name is required")
    if _is_double_quoted_token(identifier):
        return str(_parse_value(identifier))
    return identifier


def _parse_numeric_literal(token: str) -> int | float | None:
    if token.startswith(("'", '"')) and token.endswith(("'", '"')):
        return None
    try:
        return int(token)
    except ValueError:
        pass
    try:
        return float(token)
    except ValueError:
        return None


def _find_matching_parenthesis(tokens: List[str], start_index: int) -> int:
    if start_index >= len(tokens) or tokens[start_index] != "(":
        raise ValueError("Invalid SQL syntax: expected '('")

    depth = 0
    for index in range(start_index, len(tokens)):
        token = tokens[index]
        if token == "(":
            depth += 1
            continue
        if token == ")":
            depth -= 1
            if depth == 0:
                return index

    raise ValueError("Invalid SQL syntax: unmatched parenthesis")


def _find_top_level_keyword_index(tokens: List[str], keyword: str) -> int:
    depth = 0
    keyword_upper = keyword.upper()
    for index, token in enumerate(tokens):
        if token == "(":
            depth += 1
            continue
        if token == ")":
            if depth > 0:
                depth -= 1
            continue
        if depth == 0 and token.upper() == keyword_upper:
            return index
    return -1


def _tokenize_expression(text: str) -> list[str]:
    tokens: list[str] = []
    current: list[str] = []
    in_single = False
    in_double = False
    index = 0
    while index < len(text):
        char = text[index]

        if in_single:
            current.append(char)
            if char == "'":
                if index + 1 < len(text) and text[index + 1] == "'":
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_single = False
            index += 1
            continue

        if in_double:
            current.append(char)
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    current.append(text[index + 1])
                    index += 1
                else:
                    in_double = False
            index += 1
            continue

        if char.isspace():
            if current:
                tokens.append("".join(current))
                current = []
            index += 1
            continue

        if char == "|" and index + 1 < len(text) and text[index + 1] == "|":
            if current:
                tokens.append("".join(current))
                current = []
            tokens.append("||")
            index += 2
            continue

        if char in {"+", "-", "*", "/", "(", ")", ","}:
            if current:
                tokens.append("".join(current))
                current = []
            tokens.append(char)
            index += 1
            continue

        if char == "'":
            current.append(char)
            in_single = True
            index += 1
            continue

        if char == '"':
            current.append(char)
            in_double = True
            index += 1
            continue

        current.append(char)
        index += 1

    if current:
        tokens.append("".join(current))

    return _collapse_aggregate_tokens(tokens)


def _is_quoted_token(token: str) -> bool:
    return _is_single_quoted_token(token) or _is_double_quoted_token(token)


def _is_double_quoted_token(token: str) -> bool:
    return len(token) >= 2 and token.startswith('"') and token.endswith('"')


def _is_single_quoted_token(token: str) -> bool:
    return len(token) >= 2 and token.startswith("'") and token.endswith("'")


def _parse_column_identifier(token: str) -> str:
    token = token.strip()
    if _is_double_quoted_token(token):
        return token[1:-1].replace('""', '"')
    return token


def _split_qualified_identifier(token: str) -> list[str] | None:
    value = token.strip()
    if not value:
        return None

    in_double = False
    dot_index = -1
    index = 0
    while index < len(value):
        char = value[index]
        if char == '"':
            if in_double:
                if index + 1 < len(value) and value[index + 1] == '"':
                    index += 2
                    continue
                in_double = False
                index += 1
                continue
            in_double = True
            index += 1
            continue

        if char == "." and not in_double:
            if dot_index >= 0:
                return None
            dot_index = index

        index += 1

    if in_double or dot_index <= 0 or dot_index >= len(value) - 1:
        return None

    left = value[:dot_index].strip()
    right = value[dot_index + 1 :].strip()
    if not left or not right:
        return None
    return [left, right]


def _is_identifier_or_quoted(token: str) -> bool:
    token = token.strip()
    if _is_double_quoted_token(token):
        return True
    return bool(re.fullmatch(_IDENTIFIER_PATTERN, token))


def _is_qualified_identifier_or_quoted(token: str) -> bool:
    token = token.strip()
    if bool(re.fullmatch(_QUALIFIED_IDENTIFIER_PATTERN, token)):
        return True
    parts = _split_qualified_identifier(token)
    if parts is None or len(parts) != 2:
        return False
    return all(_is_identifier_or_quoted(part) for part in parts)


def _collapse_aggregate_tokens(tokens: List[str]) -> List[str]:
    collapsed: List[str] = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        upper = token.upper()
        if (
            upper == "COUNT"
            and index + 4 < len(tokens)
            and tokens[index + 1] == "("
            and tokens[index + 2].upper() == "DISTINCT"
            and tokens[index + 4] == ")"
        ):
            arg = tokens[index + 3].strip()
            collapsed.append(f"{upper}(DISTINCT {arg})")
            index += 5
            continue
        if (
            upper in _AGGREGATE_FUNCTIONS
            and index + 3 < len(tokens)
            and tokens[index + 1] == "("
            and tokens[index + 3] == ")"
        ):
            arg = tokens[index + 2].strip()
            collapsed.append(f"{upper}({arg})")
            index += 4
            continue
        collapsed.append(token)
        index += 1
    return collapsed
