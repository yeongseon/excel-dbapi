from __future__ import annotations

from typing import Any, Dict, List, Optional, Union

from ..exceptions import SqlParseError, SqlSemanticError
from .tokenizer import (
    _count_unquoted_placeholders,
    _find_matching_parenthesis,
    _is_identifier_or_quoted,
    _parse_column_identifier,
    _tokenize,
)
from .select import (
    _parse_order_by_clause_tokens,
    _parse_select,
    _validate_non_negative_pagination,
)
from .validator import _query_source_references


def _strip_outer_parens(tokens: List[str]) -> List[str]:
    """Strip one layer of outer parentheses from a token list.

    If the token list is wrapped in balanced ``(`` / ``)`` brackets that
    enclose the entire list, return the inner tokens.  Otherwise return
    the list unchanged.
    """
    if len(tokens) < 2 or tokens[0] != "(" or tokens[-1] != ")":
        return tokens
    depth = 0
    for i, tok in enumerate(tokens):
        if tok == "(":
            depth += 1
        elif tok == ")":
            depth -= 1
            if depth == 0 and i < len(tokens) - 1:
                # Closing paren is not the last token, so the outer parens
                # do NOT wrap the whole list.
                return tokens
    return tokens[1:-1]


def _extract_trailing_clauses(
    tokens: List[str],
) -> tuple[
    List[str],
    list[dict[str, Any]] | None,
    Optional[Union[int, str]],
    Optional[Union[int, str]],
]:
    """Split trailing ORDER BY / LIMIT / OFFSET from a compound's last branch.

    Returns ``(branch_tokens, order_by, limit, offset)``.
    """
    order_by: list[dict[str, Any]] | None = None
    limit: Optional[Union[int, str]] = None
    offset: Optional[Union[int, str]] = None

    # Scan backwards (at depth 0) for ORDER BY / LIMIT / OFFSET.
    # We work on the raw token list and only look for top-level keywords.
    uppers = [t.upper() for t in tokens]
    depth = 0
    keyword_positions: Dict[str, int] = {}
    for i, tok in enumerate(tokens):
        if tok == "(":
            depth += 1
        elif tok == ")":
            depth -= 1
        elif depth == 0:
            u = uppers[i]
            if u == "OFFSET" and "OFFSET" not in keyword_positions:
                keyword_positions["OFFSET"] = i
            elif u == "LIMIT" and "LIMIT" not in keyword_positions:
                keyword_positions["LIMIT"] = i
            elif u == "ORDER" and i + 1 < len(uppers) and uppers[i + 1] == "BY":
                if "ORDER BY" not in keyword_positions:
                    keyword_positions["ORDER BY"] = i

    if not keyword_positions:
        return tokens, order_by, limit, offset

    # Determine the cut point (earliest trailing clause).
    cut = min(keyword_positions.values())

    # Only accept trailing clauses that come AFTER the FROM clause of the
    # last branch. A simple heuristic: the cut must come after the last
    # top-level FROM token.
    last_from = -1
    from_depth = 0
    for i, u in enumerate(uppers[:cut]):
        tok = tokens[i]
        if tok == "(":
            from_depth += 1
        elif tok == ")":
            from_depth -= 1
        elif u == "FROM" and from_depth == 0:
            last_from = i
    if last_from == -1:
        return tokens, order_by, limit, offset

    branch_tokens = tokens[:cut]
    trail_tokens = tokens[cut:]
    trail_uppers = [t.upper() for t in trail_tokens]

    idx = 0
    while idx < len(trail_tokens):
        tu = trail_uppers[idx]
        if (
            tu == "ORDER"
            and idx + 1 < len(trail_uppers)
            and trail_uppers[idx + 1] == "BY"
        ):
            idx += 2  # skip ORDER BY
            order_tokens_list: list[str] = []
            while idx < len(trail_tokens) and trail_uppers[idx] not in {
                "LIMIT",
                "OFFSET",
            }:
                order_tokens_list.append(trail_tokens[idx])
                idx += 1
            order_by = _parse_order_by_clause_tokens(order_tokens_list)
        elif tu == "LIMIT":
            idx += 1
            if idx >= len(trail_tokens):
                raise SqlParseError("Invalid LIMIT clause format")
            val = trail_tokens[idx]
            limit = int(val) if val != "?" else val
            _validate_non_negative_pagination(limit, "LIMIT")
            idx += 1
        elif tu == "OFFSET":
            idx += 1
            if idx >= len(trail_tokens):
                raise SqlParseError("Invalid OFFSET clause format")
            val = trail_tokens[idx]
            offset = int(val) if val != "?" else val
            _validate_non_negative_pagination(offset, "OFFSET")
            idx += 1
        else:
            # Unknown trailing token — not a compound-level clause, put it back.
            return tokens, None, None, None

    return branch_tokens, order_by, limit, offset


def _parse_compound(
    query: str, params: Optional[tuple[Any, ...]]
) -> Optional[Dict[str, Any]]:
    tokens = _tokenize(query.strip())
    if not tokens:
        return None

    # Accept queries starting with either SELECT or ( (parenthesized branch).
    first_upper = tokens[0].upper()
    if first_upper != "SELECT" and tokens[0] != "(":
        return None

    query_tokens: List[List[str]] = []
    operators: List[str] = []
    current_tokens: List[str] = []
    depth = 0
    index = 0
    found_compound = False

    while index < len(tokens):
        token = tokens[index]
        upper = token.upper()

        if token == "(":
            depth += 1
            current_tokens.append(token)
            index += 1
            continue
        if token == ")":
            depth -= 1
            if depth < 0:
                if not found_compound:
                    return None
                raise SqlParseError(f"Invalid SQL query format: {query}")
            current_tokens.append(token)
            index += 1
            continue

        if depth == 0 and upper in {"UNION", "INTERSECT", "EXCEPT"}:
            if not current_tokens:
                raise SqlParseError(f"Invalid SQL query format: {query}")
            query_tokens.append(current_tokens)
            current_tokens = []
            found_compound = True

            if (
                upper == "UNION"
                and index + 1 < len(tokens)
                and tokens[index + 1].upper() == "ALL"
            ):
                operators.append("UNION ALL")
                index += 2
            else:
                operators.append(upper)
                index += 1
            continue

        current_tokens.append(token)
        index += 1

    if depth != 0 and found_compound:
        raise SqlParseError(f"Invalid SQL query format: {query}")
    if not found_compound:
        return None
    if not current_tokens:
        raise SqlParseError(f"Invalid SQL query format: {query}")

    query_tokens.append(current_tokens)
    if len(query_tokens) != len(operators) + 1:
        raise SqlParseError(f"Invalid SQL query format: {query}")

    # Extract compound-level ORDER BY / LIMIT / OFFSET from the last branch.
    last_branch = query_tokens[-1]
    last_branch, compound_order_by, compound_limit, compound_offset = (
        _extract_trailing_clauses(last_branch)
    )
    query_tokens[-1] = last_branch

    # Strip outer parens from each branch (e.g. "(SELECT ... ORDER BY ...)")
    query_tokens = [_strip_outer_parens(branch) for branch in query_tokens]

    parsed_queries: List[Dict[str, Any]] = []
    param_index = 0
    total_params = len(params) if params is not None else 0
    for segment_tokens in query_tokens:
        if not segment_tokens or segment_tokens[0].upper() != "SELECT":
            raise SqlParseError("Compound queries support only SELECT subqueries")

        segment_query = " ".join(segment_tokens)
        segment_params: Optional[tuple[Any, ...]] = None
        if params is not None:
            placeholder_count = _count_unquoted_placeholders(segment_query)
            next_index = param_index + placeholder_count
            if next_index > total_params:
                raise SqlParseError("Not enough parameters for placeholders")
            segment_params = params[param_index:next_index]
            param_index = next_index

        parsed_queries.append(_parse_select(segment_query, segment_params))

    # Resolve compound-level LIMIT / OFFSET placeholders ("?").
    if params is not None:
        if compound_limit == "?":
            if param_index >= total_params:
                raise SqlParseError("Not enough parameters for LIMIT placeholder")
            compound_limit = int(params[param_index])
            param_index += 1
        if compound_offset == "?":
            if param_index >= total_params:
                raise SqlParseError("Not enough parameters for OFFSET placeholder")
            compound_offset = int(params[param_index])
            param_index += 1

    if params is not None and param_index < total_params:
        raise SqlParseError("Too many parameters for placeholders")

    _validate_non_negative_pagination(compound_limit, "LIMIT")
    _validate_non_negative_pagination(compound_offset, "OFFSET")

    result: Dict[str, Any] = {
        "action": "COMPOUND",
        "operator": operators[0],
        "operators": operators,
        "queries": parsed_queries,
    }
    if compound_order_by is not None:
        result["order_by"] = compound_order_by
    if compound_limit is not None:
        result["limit"] = compound_limit
    if compound_offset is not None:
        result["offset"] = compound_offset
    return result


def _query_references_name(query: Dict[str, Any], name: str) -> bool:
    lowered = name.casefold()
    action = str(query.get("action", "")).upper()
    if action == "COMPOUND":
        for branch in query.get("queries", []):
            if isinstance(branch, dict) and _query_references_name(branch, name):
                return True
        return False

    if action != "SELECT":
        return False

    for source in _query_source_references(query):
        if source.casefold() == lowered:
            return True
    return False


def _parse_with_query(
    query: str,
    params: Optional[tuple[Any, ...]],
) -> Dict[str, Any]:
    tokens = _tokenize(query.strip())
    if len(tokens) < 2 or tokens[0].upper() != "WITH":
        raise SqlParseError(f"Invalid SQL query format: {query}")

    index = 1
    if tokens[index].upper() == "RECURSIVE":
        raise SqlParseError("Recursive CTEs are not supported")

    ctes: List[Dict[str, Any]] = []
    cte_names: set[str] = set()
    param_index = 0
    total_params = len(params) if params is not None else 0

    while True:
        if index >= len(tokens):
            raise SqlParseError("Invalid WITH clause: missing CTE name")

        cte_name = tokens[index]
        if not _is_identifier_or_quoted(cte_name):
            raise SqlParseError(f"Invalid CTE name: {cte_name}")
        cte_name = _parse_column_identifier(cte_name)
        lowered_name = cte_name.casefold()
        if lowered_name in cte_names:
            raise SqlSemanticError(f"Duplicate CTE name: {cte_name}")
        cte_names.add(lowered_name)
        index += 1

        if index >= len(tokens) or tokens[index].upper() != "AS":
            raise SqlParseError("Invalid WITH clause: expected AS")
        index += 1

        if index >= len(tokens) or tokens[index] != "(":
            raise SqlParseError("Invalid WITH clause: expected '(' after AS")
        end_index = _find_matching_parenthesis(tokens, index)
        subquery_tokens = tokens[index + 1 : end_index]
        if not subquery_tokens:
            raise SqlParseError("Invalid WITH clause: empty CTE query")

        subquery_sql = " ".join(subquery_tokens)
        cte_params: Optional[tuple[Any, ...]] = None
        if params is not None:
            placeholder_count = _count_unquoted_placeholders(subquery_sql)
            next_param_index = param_index + placeholder_count
            if next_param_index > total_params:
                raise SqlParseError("Not enough parameters for placeholders")
            cte_params = params[param_index:next_param_index]
            param_index = next_param_index

        parsed_cte = _parse_compound(subquery_sql, cte_params)
        if parsed_cte is None:
            parsed_cte = _parse_select(subquery_sql, cte_params)

        if _query_references_name(parsed_cte, cte_name):
            raise SqlParseError("Recursive CTEs are not supported")

        ctes.append({"type": "cte", "name": cte_name, "query": parsed_cte})
        index = end_index + 1

        if index >= len(tokens):
            raise SqlParseError("Invalid WITH clause: missing main SELECT query")
        if tokens[index] == ",":
            index += 1
            continue
        break

    main_query = " ".join(tokens[index:]).strip()
    if not main_query:
        raise SqlParseError("Invalid WITH clause: missing main SELECT query")

    main_tokens = _tokenize(main_query)
    if not main_tokens:
        raise SqlParseError("Invalid WITH clause: missing main SELECT query")
    if main_tokens[0].upper() != "SELECT" and main_tokens[0] != "(":
        raise SqlParseError("WITH clause requires a SELECT query")

    remaining_params: Optional[tuple[Any, ...]] = None
    if params is not None:
        remaining_params = params[param_index:]

    parsed = _parse_compound(main_query, remaining_params)
    if parsed is None:
        if main_tokens[0] == "(":
            raise SqlParseError(f"Unsupported SQL action: {main_tokens[0]}")
        parsed = _parse_select(main_query, remaining_params)

    parsed["ctes"] = ctes
    return parsed
