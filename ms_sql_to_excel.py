import pyodbc
import pandas as pd
import os
import argparse
from datetime import datetime

def main(invoice_date):
    # Parse the input dates
    try:
        invoice_date = datetime.strptime(invoice_date, "%Y-%m-%d")
    except ValueError as e:
        print(f"Error parsing dates: {e}")
        return

    # Print the dates (for debugging purposes)
    print(f"Invoice Date: {invoice_date}")

    # Use environment variables for sensitive information
    server = os.getenv('AZURE_SQL_SERVER')
    database = os.getenv('AZURE_SQL_DATABASE')
    username = os.getenv('AZURE_SQL_USERNAME')
    password = os.getenv('AZURE_SQL_PASSWORD')
    driver = '{ODBC Driver 17 for SQL Server}'

    # Print the environment variables
    print("Here are the environment variables we are going to use:")
    print("server: ", server)
    print("database: ", database)
    print("username: ", username)
    print("password: ", password)
    print("driver: ", driver)
    print("")


    # Connection string
    connection_string = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'

    # Print the environment variables
    print("Here is the connection string:")
    print("connection_string: ", connection_string)
    print("")

    try:
        # Establish the connection
        conn = pyodbc.connect(connection_string)
        print("Connection established successfully.")

        # Read the contents of the SQL files into variables
        with open('Journals2Expenses.sql', 'r') as file:
            sql_query_j2e = file.read()

        with open('SalesReceipts.sql', 'r') as file:
            sql_query_sr = file.read()

        # Replace INVOICE_DATE in the SQL queries
        sql_query_j2e = sql_query_j2e.replace('INVOICE_DATE', invoice_date.strftime("%Y-%m-%d %H:%M:%S"))
        sql_query_sr = sql_query_sr.replace('INVOICE_DATE', invoice_date.strftime("%Y-%m-%d %H:%M:%S"))

        # Queries for the tables
        queries = {
            'Journals to Expenses': sql_query_j2e,
            'Sales_Receipts': sql_query_sr
        }

        # Load data from SQL Database into DataFrames
        data_frames = {}
        for table_name, query in queries.items():
            data_frames[table_name] = pd.read_sql(query, conn)

        # CSVs for the conversions
        csvs = {
            'AP_Categories': "APCategories.csv"
        }

        # Load data from CSV files into separate DataFrames
        data_frames2 = {}
        for table_name, csv in csvs.items():
            data_frames2[table_name] = pd.read_csv(csv)

        # Assign Expense Account on Journals to Expenses sheet based on AP Categories DataFrame
        merged = pd.merge(data_frames['Journals to Expenses'], data_frames2['AP_Categories'], on='ItemCategoryNumber', how='left')
        data_frames['Journals to Expenses']['Expense Account'] = merged['GLCode']

        # Remove fields not needed from data frames
        data_frames['Journals to Expenses'] = data_frames['Journals to Expenses'].drop(columns=['ItemCategoryNumber'])
        data_frames['Journals to Expenses'] = data_frames['Journals to Expenses'].drop(columns=['ItemCategoryName'])

        # Format the invoice date
        invoice_date_str = invoice_date.strftime("%Y%m%d")

        # Prefix the DataFrame name with the formatted invoice date
        new_df_name = f"{invoice_date_str}-Journals to Expenses"
        data_frames[new_df_name] = data_frames.pop('Journals to Expenses')

        # Export each DataFrame to a separate Excel sheet
        with pd.ExcelWriter('INVOICE REPORT.xlsx') as writer:
            for table_name, df in data_frames.items():
                df.to_excel(writer, sheet_name=table_name, index=False)

        print("Data has been exported to INVOICE REPORT.xlsx successfully.")

    except pyodbc.Error as e:
        print(f"Error: {e}")
    finally:
        # Close the connection
        if 'conn' in locals():
            conn.close()
            print("Connection closed.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="A script that takes two date parameters.")
    parser.add_argument('invoice_date', type=str, help='The start date in YYYY-MM-DD format')
    args = parser.parse_args()
    main(args.invoice_date)