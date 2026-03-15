from excel_dbapi.connection import ExcelConnection


with ExcelConnection("tests/data/sample.xlsx", autocommit=False) as conn:
    cur = conn.cursor()
    cur.execute("INSERT INTO Sheet1 (id, name) VALUES (?, ?)", (101, "Student"))
    conn.commit()
