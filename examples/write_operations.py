# examples/write_operations.py

from excel_dbapi.connection import ExcelConnection


def main():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Sheet1 (id, name) VALUES (10, 'Zoe')")
        cursor.execute("UPDATE Sheet1 SET name = 'Zoey' WHERE id = 10")
        cursor.execute("DELETE FROM Sheet1 WHERE id = 10")

        cursor.executemany(
            "INSERT INTO Sheet1 (id, name) VALUES (?, ?)",
            [(11, "Mina"), (12, "Noah")],
        )


if __name__ == "__main__":
    main()
