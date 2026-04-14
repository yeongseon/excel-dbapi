"""Tests for subqueries in UPDATE/DELETE WHERE clauses."""

from pathlib import Path

from openpyxl import Workbook

from excel_dbapi.connection import ExcelConnection


def _create_test_workbook(path: Path) -> None:
    workbook = Workbook()

    employees = workbook.active
    assert employees is not None
    employees.title = "employees"
    employees.append(["id", "name", "dept_id", "salary"])
    employees.append([1, "Alice", 10, 50000])
    employees.append([2, "Bob", 10, 60000])
    employees.append([3, "Carol", 20, 55000])
    employees.append([4, "Dan", 30, 45000])
    employees.append([5, "Eve", 99, 70000])

    departments = workbook.create_sheet("departments")
    departments.append(["id", "name", "active"])
    departments.append([10, "Engineering", 1])
    departments.append([20, "Sales", 1])
    departments.append([30, "HR", 0])

    workbook.save(path)


def _run(path: Path, query: str) -> int:
    """Execute a DML statement and return rowcount."""
    with ExcelConnection(str(path), engine="openpyxl", autocommit=True) as conn:
        cursor = conn.cursor()
        cursor.execute(query)
        return cursor.rowcount


def _select(path: Path, query: str) -> list[tuple[object, ...]]:
    with ExcelConnection(str(path), engine="openpyxl") as conn:
        cursor = conn.cursor()
        cursor.execute(query)
        return cursor.fetchall()


def test_update_with_in_subquery(tmp_path: Path) -> None:
    file_path = tmp_path / "subq_update_in.xlsx"
    _create_test_workbook(file_path)

    rowcount = _run(
        file_path,
        "UPDATE employees SET salary = 99999 "
        "WHERE dept_id IN (SELECT id FROM departments WHERE name = 'Engineering')",
    )
    assert rowcount == 2

    rows = _select(file_path, "SELECT name, salary FROM employees ORDER BY name")
    assert rows == [
        ("Alice", 99999),
        ("Bob", 99999),
        ("Carol", 55000),
        ("Dan", 45000),
        ("Eve", 70000),
    ]


def test_delete_with_not_in_subquery(tmp_path: Path) -> None:
    file_path = tmp_path / "subq_delete_not_in.xlsx"
    _create_test_workbook(file_path)

    rowcount = _run(
        file_path,
        "DELETE FROM employees "
        "WHERE dept_id NOT IN (SELECT id FROM departments WHERE active = 1)",
    )
    # dept_id 30 (active=0) and 99 (not in departments at all) should be deleted
    assert rowcount == 2

    rows = _select(file_path, "SELECT name FROM employees ORDER BY name")
    assert rows == [("Alice",), ("Bob",), ("Carol",)]


def test_delete_with_in_subquery(tmp_path: Path) -> None:
    file_path = tmp_path / "subq_delete_in.xlsx"
    _create_test_workbook(file_path)

    rowcount = _run(
        file_path,
        "DELETE FROM employees "
        "WHERE dept_id IN (SELECT id FROM departments WHERE active = 0)",
    )
    assert rowcount == 1  # Dan (dept_id=30, HR inactive)

    rows = _select(file_path, "SELECT name FROM employees ORDER BY name")
    assert rows == [("Alice",), ("Bob",), ("Carol",), ("Eve",)]


def test_update_with_not_in_subquery(tmp_path: Path) -> None:
    file_path = tmp_path / "subq_update_not_in.xlsx"
    _create_test_workbook(file_path)

    rowcount = _run(
        file_path,
        "UPDATE employees SET salary = 0 "
        "WHERE dept_id NOT IN (SELECT id FROM departments)",
    )
    assert rowcount == 1  # Eve (dept_id=99 not in departments)

    rows = _select(file_path, "SELECT name, salary FROM employees WHERE salary = 0")
    assert rows == [("Eve", 0)]


def test_update_subquery_with_where_in_subquery(tmp_path: Path) -> None:
    """Subquery itself has a WHERE clause."""
    file_path = tmp_path / "subq_update_nested_where.xlsx"
    _create_test_workbook(file_path)

    rowcount = _run(
        file_path,
        "UPDATE employees SET salary = 77777 "
        "WHERE dept_id IN (SELECT id FROM departments WHERE name = 'Sales')",
    )
    assert rowcount == 1  # Carol (dept_id=20)

    rows = _select(file_path, "SELECT name, salary FROM employees WHERE name = 'Carol'")
    assert rows == [("Carol", 77777)]


def test_update_subquery_no_matches(tmp_path: Path) -> None:
    """Subquery returns empty result set."""
    file_path = tmp_path / "subq_update_empty.xlsx"
    _create_test_workbook(file_path)

    rowcount = _run(
        file_path,
        "UPDATE employees SET salary = 0 "
        "WHERE dept_id IN (SELECT id FROM departments WHERE name = 'NonExistent')",
    )
    assert rowcount == 0

    # All salaries unchanged
    rows = _select(file_path, "SELECT salary FROM employees ORDER BY salary")
    assert rows == [(45000,), (50000,), (55000,), (60000,), (70000,)]


def test_delete_subquery_no_matches(tmp_path: Path) -> None:
    """Subquery returns empty result set - no rows deleted."""
    file_path = tmp_path / "subq_delete_empty.xlsx"
    _create_test_workbook(file_path)

    rowcount = _run(
        file_path,
        "DELETE FROM employees "
        "WHERE dept_id IN (SELECT id FROM departments WHERE name = 'NonExistent')",
    )
    assert rowcount == 0

    rows = _select(file_path, "SELECT COUNT(*) FROM employees")
    assert rows == [(5,)]
