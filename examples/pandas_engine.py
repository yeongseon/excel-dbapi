# examples/pandas_engine.py

from excel_dbapi.connection import ExcelConnection


def main():
    with ExcelConnection("tests/data/sample.xlsx", engine="pandas") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        print(cursor.fetchall())


if __name__ == "__main__":
    main()
