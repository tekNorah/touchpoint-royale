import pyodbc  # !pip install pyodbc
import pandas as pd  # !pip install pandas
import openpyxl  # !pip install openpyxl

# Connection details
server = 'pointdbserver.database.windows.net'
database = 'NCR-MothersEntJA'
username = 'Norah'  # replace with your username
password = '!WSX#Rrrr45'  # replace with your password
driver = '{ODBC Driver 18 for SQL Server}'

# Establish the connection
connection_string = f"""
DRIVER={driver};
SERVER={server};
PORT=1433;
DATABASE={database};
UID={username};
PWD={password}
"""

conn = pyodbc.connect(connection_string)

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

# Close the connection
conn.close()