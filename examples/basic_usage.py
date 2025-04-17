# examples/basic_usage.py

from excel_dbapi.connection import ExcelConnection

def main():
    # Open Excel file using ExcelConnection
    with ExcelConnection("tests/data/sample.xlsx") as conn:
        # Create a cursor
        cursor = conn.cursor()
        
        # Execute a SQL-like query
        cursor.execute("SELECT * FROM Sheet1")
        
        # Fetch all results
        results = cursor.fetchall()
        
        # Print the results
        for row in results:
            print(row)

if __name__ == "__main__":
    main()
