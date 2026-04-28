from __future__ import annotations

import argparse
from collections.abc import Sequence
from pathlib import Path
import sys

from excel_dbapi import Error, connect
from excel_dbapi.engines.result import ExecutionResult
from excel_dbapi.reflection import list_tables

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excel-dbapi")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect", help="Show workbook summary")
    _ = inspect_parser.add_argument("file_path", help="Path to workbook")
    _ = inspect_parser.add_argument(
        "--engine",
        default=None,
        help="Engine backend to use (default: auto-detect)",
    )

    tables_parser = subparsers.add_parser("tables", help="List workbook sheet names")
    _ = tables_parser.add_argument("file_path", help="Path to workbook")
    _ = tables_parser.add_argument(
        "--engine",
        default=None,
        help="Engine backend to use (default: auto-detect)",
    )

    schema_parser = subparsers.add_parser(
        "schema", help="Show sheet headers and row counts"
    )
    _ = schema_parser.add_argument("file_path", help="Path to workbook")
    _ = schema_parser.add_argument(
        "sheet", nargs="?", default=None, help="Sheet name (default: all sheets)"
    )
    _ = schema_parser.add_argument(
        "--engine",
        default=None,
        help="Engine backend to use (default: auto-detect)",
    )

    query_parser = subparsers.add_parser("query", help="Execute SQL query")
    _ = query_parser.add_argument("file_path", help="Path to workbook")
    _ = query_parser.add_argument("sql", help="SQL query string")
    _ = query_parser.add_argument(
        "--engine",
        default=None,
        help="Engine backend to use (default: auto-detect)",
    )
    _ = query_parser.add_argument(
        "--data-only",
        action="store_true",
        default=False,
        dest="data_only",
        help="Open workbook with data_only=True (replaces formulas with cached values)",
    )

    return parser


def _headers_text(headers: list[str]) -> str:
    return ", ".join(headers) if headers else "(none)"


def _print_inspect(file_path: str, engine: str | None) -> int:
    with connect(file_path, engine=engine) as conn:
        workbook_name = Path(file_path).name
        engine_name = conn.engine_name
        print(f"Workbook: {workbook_name}")
        print(f"Engine: {engine_name}")
        print()
        print("Sheets:")
        for sheet_name in list_tables(conn):
            table = conn.engine.read_sheet(sheet_name)
            print(f"  - {sheet_name}")
            print(f"    rows: {len(table.rows)}")
            print(f"    columns: {len(table.headers)}")
            print(f"    headers: {_headers_text(table.headers)}")
    return 0


def _print_tables(file_path: str, engine: str | None) -> int:
    with connect(file_path, engine=engine) as conn:
        for sheet_name in list_tables(conn):
            print(sheet_name)
    return 0


def _print_schema(file_path: str, engine: str | None, sheet: str | None = None) -> int:
    with connect(file_path, engine=engine) as conn:
        sheets = [sheet] if sheet else list_tables(conn)
        for sheet_name in sheets:
            table = conn.engine.read_sheet(sheet_name)
            print(f"{sheet_name}:")
            print(f"  rows: {len(table.rows)}")
            print(f"  headers: {_headers_text(table.headers)}")
    return 0


def _stringify_cell(value: object) -> str:
    return "" if value is None else str(value)


def _format_results(headers: list[str], rows: list[tuple[object, ...]]) -> str:
    rendered_rows: list[list[str]] = [
        [_stringify_cell(cell) for cell in row] for row in rows
    ]
    table_data: list[list[str]] = [headers, *rendered_rows]
    widths = [max(len(line[i]) for line in table_data) for i in range(len(headers))]
    separator = "-+-".join("-" * width for width in widths)

    def _format_row(cells: Sequence[str]) -> str:
        return " | ".join(cell.ljust(widths[i]) for i, cell in enumerate(cells))

    formatted_lines = [_format_row(headers), separator]
    formatted_lines.extend(_format_row(row) for row in rendered_rows)
    return "\n".join(formatted_lines)


def _description_to_headers(result: ExecutionResult) -> list[str]:
    if result.description:
        return [col[0] if col[0] is not None else "" for col in result.description]
    if result.rows:
        return [f"col{i + 1}" for i in range(len(result.rows[0]))]
    return []


def _print_query(file_path: str, sql: str, engine: str | None, *, data_only: bool = False) -> int:
    with connect(file_path, engine=engine, data_only=data_only) as conn:
        result = conn.execute(sql)

    headers = _description_to_headers(result)
    is_query = result.action.upper() in ("SELECT", "COMPOUND")

    if is_query:
        rows = [tuple(row) for row in result.rows]
        if headers:
            print(_format_results(headers, rows))
        print(f"{len(rows)} row(s)")
    else:
        print(f"OK ({result.rowcount} rows affected)")
    return 0


def _run(args: argparse.Namespace) -> int:
    command_obj = getattr(args, "command", None)
    file_path_obj = getattr(args, "file_path", None)
    engine_obj = getattr(args, "engine", None)

    if not isinstance(command_obj, str):
        raise ValueError("Missing command")
    if not isinstance(file_path_obj, str):
        raise ValueError("Missing workbook file path")

    command = command_obj
    file_path = file_path_obj
    engine: str | None = engine_obj if isinstance(engine_obj, str) else None

    if command == "inspect":
        return _print_inspect(file_path, engine)
    if command == "tables":
        return _print_tables(file_path, engine)
    if command == "schema":
        sheet_obj = getattr(args, "sheet", None)
        sheet: str | None = sheet_obj if isinstance(sheet_obj, str) else None
        return _print_schema(file_path, engine, sheet)
    if command == "query":
        sql_obj = getattr(args, "sql", None)
        if not isinstance(sql_obj, str):
            raise ValueError("Missing SQL query")
        sql = sql_obj
        data_only: bool = getattr(args, "data_only", False)
        return _print_query(file_path, sql, engine, data_only=data_only)

    raise ValueError(f"Unknown command: {command}")

def main(argv: Sequence[str] | None = None) -> None:
    parser = _build_parser()
    parsed_args = parser.parse_args(list(argv) if argv is not None else None)
    try:
        code = _run(parsed_args)
    except Error as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
    except FileNotFoundError as exc:
        print(f"Error: file not found: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
    raise SystemExit(code)


if __name__ == "__main__":
    main()
