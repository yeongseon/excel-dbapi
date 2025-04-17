from excel_dbapi.connection import ExcelConnection


def test_full_flow():
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Sheet1")
        results = cursor.fetchall()

        assert isinstance(results, list)
        assert isinstance(results[0], dict)
        assert "name" in results[0]  # assuming sample.xlsx has a 'name' column
