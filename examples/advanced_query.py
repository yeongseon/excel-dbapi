# examples/advanced_query.py

from excel_dbapi.connection import ExcelConnection


def main():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, name FROM Sheet1 WHERE id >= ? AND name != ? ORDER BY id DESC LIMIT 2",
            (1, "Alice"),
        )
        print(cursor.fetchall())


if __name__ == "__main__":
    main()
