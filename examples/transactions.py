# examples/transactions.py

from excel_dbapi.connection import ExcelConnection


def main():
    with ExcelConnection("tests/data/sample.xlsx", autocommit=False) as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE Sheet1 SET name = 'Ann' WHERE id = 1")
        conn.rollback()


if __name__ == "__main__":
    main()
