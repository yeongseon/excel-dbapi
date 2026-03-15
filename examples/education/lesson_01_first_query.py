from excel_dbapi.connection import ExcelConnection


with ExcelConnection("tests/data/sample.xlsx") as conn:
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM Sheet1 ORDER BY id LIMIT 5")
    for row in cur.fetchall():
        print(row)
