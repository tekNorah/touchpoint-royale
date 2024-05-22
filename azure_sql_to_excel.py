import pyodbc
import pandas as pd
import openpyxl
import os


# Use environment variables for sensitive information
server = os.getenv('AZURE_SQL_SERVER')
database = os.getenv('AZURE_SQL_DATABASE')
username = os.getenv('AZURE_SQL_USERNAME')
password = os.getenv('AZURE_SQL_PASSWORD')
driver = '{ODBC Driver 18 for SQL Server}'

# Print the environment variables
print("Here are the environment variables we are going to use:")
print("server: ",server)
print("database: ",database)
print("username: ",username)
print("password: ",password)
print("driver: ",driver)
print("")


# Connection string
connection_string = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Print the environment variables
print("Here is the connection string:")
print("connection_string: ",connection_string)
print("")

try:
    # Establish the connection
    conn = pyodbc.connect(connection_string)
    print("Connection established successfully.")

    # Queries for the tables
    queries = {
        'ST_Sites': 'SELECT * FROM ST_Sites',
        'ST_Deposits': 'SELECT * FROM ST_Deposits',
        'ST_Register': 'SELECT * FROM ST_Register'
    }

    # Load data into DataFrames
    data_frames = {}
    for table_name, query in queries.items():
        data_frames[table_name] = pd.read_sql(query, conn)

    # Export each DataFrame to a separate Excel sheet
    with pd.ExcelWriter('output.xlsx') as writer:
        for table_name, df in data_frames.items():
            df.to_excel(writer, sheet_name=table_name, index=False)

    print("Data has been exported to output.xlsx successfully.")

except pyodbc.Error as e:
    print(f"Error: {e}")
finally:
    # Close the connection
    if 'conn' in locals():
        conn.close()
        print("Connection closed.")