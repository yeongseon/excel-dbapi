# examples/basic_usage.py

from excel_dbapi.connection import ExcelConnection


def main():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        for row in cursor.fetchall():
            print(row)


if __name__ == "__main__":
    main()
