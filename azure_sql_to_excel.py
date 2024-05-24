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
        'Batch_Header': """SELECT
    StoreID,
    RIGHT('00000'+CAST(ROW_NUMBER() OVER(ORDER BY StoreID, SalesDate) AS VARCHAR(5)),5) AS ENTRYNO,
    0 AS ENTRYTYPE,
    MAX(BagNumber) AS REFERENCE,
    RIGHT('00'+CAST(MONTH(SalesDate) AS VARCHAR(2)),2) as PERIOD,
    convert(varchar,SalesDate,101) AS DATE,
    NULL AS DATECHQPRN,
    'FALSE' AS SWCHQPRN,
    NULL AS MISCCODE,
    NULL AS TEXTDESC,
    NULL AS DISTCODE,
    RIGHT('00'+CAST(1 AS VARCHAR(2)),2) AS CHARGECODE,
    0 AS CHRGAMOUNT,
    NULL AS NODETAILS,
    SUM(Amount) AS TOTAMOUNT,
    0 AS TOTTAX,
    0 AS TAXPERCNT,
    'JAD' AS BK2GLCURHM,
    'SP' AS BK2GLRTTYP,
    'JAD' AS BK2GLCURSR,
    convert(varchar,SalesDate,101) AS BK2GLDATE,
    1 AS BK2GLRATE,
    0 AS BK2GLSPRD,
    1 AS BK2GLOP,
    1 AS BK2GLDTMTH,
    'JAD' AS BT2GLCURHM,
    'SP' AS BT2GLRTTYP,
    'JAD' AS BT2GLCURSR,
    convert(varchar,SalesDate,101) AS BT2GLDATE,
    1 AS BT2GLRATE,
    0 AS BT2GLSPRD,
    1 AS BT2GLOP,
    1 AS BT2GLDTMTH,
    'JAD' AS MS2GLCURHM,
    'SP' AS MS2GLRTTYP,
    'JAD' AS MS2GLCURSR,
    convert(varchar,SalesDate,101) AS MS2GLDATE,
    1 AS MS2GLRATE,
    0 AS MS2GLSPRD,
    1 AS MS2GLOP,
    1 AS MS2GLDTMTH,
    'FALSE' AS SWCASH,
    2 AS BTCHNODEC,
    2 AS MISCNODEC,
    NULL AS TAXGROUP,
    NULL AS CUSTCHQNO,
    0 AS NOSUBDETL,
    0 AS APPLAMOUNT,
	0 AS APPLDISC,
	NULL AS ACCTNAT,
	0 AS ADJAMOUNT,
	NULL AS PROFILEID,
	'FALSE' AS SWINTERCO,
	YEAR(SalesDate) AS FISCYR,
	NULL AS CCTYPE,
	'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=' AS CCNUMBER,
	NULL AS CCNAME,
	NULL AS CCEXP,
	NULL AS CCAUTHCODE,
	NULL AS XCCNUMBER,
	0 AS SERIAL,
	SUM(Amount) AS BANKAMOUNT,
	SUM(Amount) AS BTCHAMOUNT,
	SUM(Amount) AS MISCAMOUNT,
	SUM(Amount) AS FUNCAMOUNT,
	NULL AS HDRDEBIT,
    NULL AS HDRCREDIT,
	NULL AS TAXAUTH1,
	NULL AS TAXAUTH2,
	NULL AS TAXAUTH3,
	NULL AS TAXAUTH4,
	NULL AS TAXAUTH5,
	NULL AS TXAU1DESC,
	NULL AS TXAU2DESC,
	NULL AS TXAU3DESC,
	NULL AS TXAU4DESC,
	NULL AS TXAU5DESC,
	1 AS TAXCLASS1,
	1 AS TAXCLASS2,
	1 AS TAXCLASS3,
	1 AS TAXCLASS4,
	1 AS TAXCLASS5,
	0 AS BASETAX1,
	0 AS BASETAX2,
	0 AS BASETAX3,
	0 AS BASETAX4,
	0 AS BASETAX5,
	0 AS AMTTAX1,
	0 AS AMTTAX2,
	0 AS AMTTAX3,
	0 AS AMTTAX4,
	0 AS AMTTAX5,
	0 AS SWTAXINCL1,
	0 AS SWTAXINCL2,
	0 AS SWTAXINCL3,
	0 AS SWTAXINCL4,
	0 AS SWTAXINCL5,
	RIGHT('00'+CAST(1 AS VARCHAR(2)),2) AS BANKCODE,
	'TRUE' AS SWPOSTED,
	0 AS 'VALUES',
	0 AS PROCESSCMD,
	0 AS TOTUNAPPL,
	0 AS TOTAPPLAMT,
	0 AS TOTAPPLDSC,
	0 AS ALLOCMODE,
	0 AS ALLOCAMT,
	2 AS CLASSTYPE,
	1 AS CLASSAXIS,
	0 AS DATALEVEL,
	0 AS RECXCNTER,
	1 AS RATERC,
	'SP' AS RATETYPERC,
	1 AS RATEOPRC,
    convert(varchar,SalesDate,101) AS RATEDATERC,
    'JAD' AS CODECURNRC,
    0 AS TXAMT1RC,
	0 AS TXAMT2RC,
	0 AS TXAMT3RC,
	0 AS TXAMT4RC,
	0 AS TXAMT5RC,
	0 AS TXTOTRC,
	0 AS TXALLRC,
	0 AS TXEXPRC,
	0 AS TXRECRC,
	0 AS AMTRECTAX,
	0 AS AMTEXPTAX,
	0 AS TXBSE1HC,
	0 AS TXBSE2HC,
	0 AS TXBSE3HC,
	0 AS TXBSE4HC,
	0 AS TXBSE5HC,
	0 AS TXAMT1HC,
	0 AS TXAMT2HC,
	0 AS TXAMT3HC,
	0 AS TXAMT4HC,
	0 AS TXAMT5HC,
	NULL ACCTREC1,
	NULL AS ACCTREC2,
	NULL AS ACCTREC3,
	NULL AS ACCTREC4,
	NULL AS ACCTREC5,
	NULL AS ACCTEXP1,
	NULL AS ACCTEXP2,
	NULL AS ACCTEXP3,
	NULL AS ACCTEXP4,
	NULL AS ACCTEXP5,
	0 AS TXEXCLTC,
	0 AS TXEXCLHC,
	0 AS TXEXCLBC,
	0 AS TXEXCLMC,
	0 AS TXINCLTC,
	0 AS TXINCLHC,
	0 AS TXINCLBC,
	0 AS TXINCLMC,
	NULL ARAPBATCH,
	NULL AS ARAPENTRY,
	'FALSE' AS SWCHEQUE,
	'TRUE' AS SWEFT,
	0 AS RXMTCHSEQ,
	NULL AS RXTRNSCODE,
	NULL AS RXCATEGORY,
	0 AS REVUNIQ,
	0 AS NEWREVUNIQ,
	'TASHI' AS ENTEREDBY,
	NULL AS WORKDESC,
	0 AS BTYPE,
	NULL AS MSCURR,
	0 AS BALDUE,
	'JAD' AS BKCURR,
	0 AS SWGLQTY,
	0 AS GLQTYDEC,
	241175188 AS CBBCTLRVH,
	241175188 AS CBBCTLVW,
	241165428 AS CBBTDTRVH,
	241165428 AS CBBTDTVW,
	0 AS CBERRNUM,
	NULL AS CBERRDESC,
	NULL AS CBERRFUNC,
	'FALSE' AS NOVALTXGRP
FROM ST_Deposits
WHERE SalesDate = '2024-05-01 00:00:00.000'
GROUP BY StoreID, SalesDate""",
        'Batch_Detail': 'SELECT * FROM ST_Register'
    }

    # Load data into DataFrames
    data_frames = {}
    for table_name, query in queries.items():
        data_frames[table_name] = pd.read_sql(query, conn)

    # Export each DataFrame to a separate Excel sheet
    with pd.ExcelWriter('CASH REPORT.xlsx') as writer:
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