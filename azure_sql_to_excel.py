import random

import pyodbc
import pandas as pd
import os


# Use environment variables for sensitive information
server = os.getenv('AZURE_SQL_SERVER')
database = os.getenv('AZURE_SQL_DATABASE')
username = os.getenv('AZURE_SQL_USERNAME')
password = os.getenv('AZURE_SQL_PASSWORD')
driver = '{ODBC Driver 18 for SQL Server}'

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

    # Queries for the tables
    # noinspection SqlNoDataSourceInspection
    queries = {
        'Batch_Header': """
            SELECT
                CONVERT(int, StoreID) AS StoreID,
                NULL BATCHID,
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
        'Batch_Detail': """
            SELECT
                StoreID,
                DATE,
                BATCHID,
                ENTRYNO,
                DETAILNO,
                SRCECODE,
                TEXTDESC,
                ACCTID,
                TAXCODE,
                TAXTYPE,
                TAXPERCNT,
                TAXAMOUNT,
                DTLAMOUNT,
                QUANTITY,
                COMMENTS,
                RCPTNO,
                SWCASH,
                RCPTDESC,
                MISCCODE,
                MISCBKLINE,
                RCPTENTRY,
                ACCTIDUF,
                ALLOCMODE,
                CASE ACCTDESC
                    WHEN 'Deposit1' THEN 'SALES CLEARING ACCOUNT'
                    WHEN 'Deposit2' THEN 'RECEIVABLE - DELIVERY PARTNERS'
                    WHEN 'Deposit3' THEN 'CATERING REC.'
                    WHEN 'Deposit4' THEN 'LUNCH VOUCHER'
                    WHEN 'Deposit5' THEN 'CREDIT CARD SALES REC.'
                    WHEN 'Deposit6' THEN 'CREDIT WHOLESALE PATTY'
                    WHEN 'Deposit7' THEN 'CASH SHORTAGE/OVERAGE'
                    WHEN 'Deposit8' THEN 'BNS USD CLEARING A/C'
                    ELSE ACCTDESC END AS ACCTDESC,
                0 AS ACCTQTYSW,
                ACCTTAX,
                ACCTTAXUF,
                TAXDESC,
                TAXQTYSW,
                ADJAMOUNT,
                SWJOB,
                DEBITAMT,
                CREDITAMT,
                SWALLOCATE,
                PJCAMT,
                PJCDISC,
                ENTRYTYPE,
                DOCNUMBER,
                TOTTAX,
                SWMANLTX,
                BASETAX1,
                BASETAX2,
                BASETAX3,
                BASETAX4,
                BASETAX5,
                TAXCLASS1,
                TAXCLASS2,
                TAXCLASS3,
                TAXCLASS4,
                TAXCLASS5,
                SWTAXINCL1,
                SWTAXINCL2,
                SWTAXINCL3,
                SWTAXINCL4,
                SWTAXINCL5,
                RATETAX1,
                RATETAX2,
                RATETAX3,
                RATETAX4,
                RATETAX5,
                AMTTAX1,
                AMTTAX2,
                AMTTAX3,
                AMTTAX4,
                AMTTAX5,
                MISCTAX1,
                MISCTAX2,
                MISCTAX3,
                MISCTAX4,
                MISCTAX5,
                GLTAXAMT1,
                GLTAXAMT2,
                GLTAXAMT3,
                GLTAXAMT4,
                GLTAXAMT5,
                BKTAXAMT1,
                BKTAXAMT2,
                BKTAXAMT3,
                BKTAXAMT4,
                BKTAXAMT5,
                TCAMTINCTX,
                GLAMTINCTX,
                BKAMTINCTX,
                MSAMTINCTX,
                MISCAMOUNT,
                GLAMOUNT,
                BKAMOUNT,
                TOTAPPLAMT,
                TOTAPPLDSC,
                TOTUNAPPL,
                PJCCOST,
                NOSUBDETL,
                'VALUES',
                PROCESSCMD,
                AMTTAXREC1,
                AMTTAXREC2,
                AMTTAXREC3,
                AMTTAXREC4,
                AMTTAXREC5,
                AMTTAXEXP1,
                AMTTAXEXP2,
                AMTTAXEXP3,
                AMTTAXEXP4,
                AMTTAXEXP5,
                AMTTAXTOBE,
                TXAMT1RC,
                TXAMT2RC,
                TXAMT3RC,
                TXAMT4RC,
                TXAMT5RC,
                TXTOTRC,
                TXALLRC,
                TXEXP1RC,
                TXEXP2RC,
                TXEXP3RC,
                TXEXP4RC,
                TXEXP5RC,
                TXREC1RC,
                TXREC2RC,
                TXREC3RC,
                TXREC4RC,
                TXREC5RC,
                TXBSE1HC,
                TXBSE2HC,
                TXBSE3HC,
                TXBSE4HC,
                TXBSE5HC,
                TXREC1HC,
                TXREC2HC,
                TXREC3HC,
                TXREC4HC,
                TXREC5HC,
                TXEXP1HC,
                TXEXP2HC,
                TXEXP3HC,
                TXEXP4HC,
                TXEXP5HC,
                TXALLHC,
                TXALL1HC,
                TXALL2HC,
                TXALL3HC,
                TXALL4HC,
                TXALL5HC,
                TXALL1TC,
                TXALL2TC,
                TXALL3TC,
                TXALL4TC,
                TXALL5TC,
                TXEXCLTC,
                TXEXCLHC,
                TXEXCLBC,
                TXEXCLMC,
                REVUNIQ,
                NEWREVUNIQ,
                RVDETAILNO,
                SWNODEFS,
                CBBTHDRVH,
                CBBTHDVW,
                CBBTSDRVH,
                CBBTSDVW,
                CBERRNUM,
                CBERRDESC,
                CBERRFUNC,
                SWNOTAXVAL,
                SWREFUND,
                SWZEROENT,
                SWAUTALLOC,
                SWINCLDISC,
                SWIGNORE
            FROM
               (SELECT
                    CONVERT(int, StoreID) AS StoreID,
                    convert(varchar,SalesDate,101) as 'DATE',
                    NULL AS BATCHID,
                    RIGHT('00000'+CAST(ROW_NUMBER() OVER(ORDER BY StoreID, SalesDate) AS VARCHAR(5)),5) AS ENTRYNO,
                    NULL AS DETAILNO,
                    'LODG' AS SRCECODE,
                    NULL AS TEXTDESC,
                    NULL AS ACCTID,
                    RIGHT('01'+CAST(1 AS VARCHAR(2)),2) AS TAXCODE,
                    0 AS TAXTYPE,
                    0 AS TAXPERCNT,
                    0 AS TAXAMOUNT,
                    0 AS 'QUANTITY',
                    NULL AS COMMENTS,
                    NULL AS RCPTNO,
                    'FALSE' AS SWCASH,
                    NULL AS RCPTDESC,
                    NULL AS MISCCODE,
                    0 AS MISCBKLINE,
                    NULL AS RCPTENTRY,
                    NULL AS ACCTIDUF,
                    0 AS ALLOCMODE,
                    Deposit1, Deposit2, Deposit3, Deposit4, Deposit5, Deposit6, Deposit7, Deposit8,
                    0 AS ACCTQTYSW,
                    NULL AS ACCTTAX,
                    NULL AS ACCTTAXUF,
                    NULL AS TAXDESC,
                    0 AS TAXQTYSW,
                    0 AS ADJAMOUNT,
                    'FALSE' AS SWJOB,
                    NULL AS DEBITAMT,
                    NULL AS CREDITAMT,
                    'FALSE' AS SWALLOCATE,
                    0 AS PJCAMT,
                    0 AS PJCDISC,
                    0 AS ENTRYTYPE,
                    NULL AS DOCNUMBER,
                    0 AS TOTTAX,
                    0 AS SWMANLTX,
                    0 AS BASETAX1,
                    0 AS BASETAX2,
                    0 AS BASETAX3,
                    0 AS BASETAX4,
                    0 AS BASETAX5,
                    1 AS TAXCLASS1,
                    1 AS TAXCLASS2,
                    1 AS TAXCLASS3,
                    1 AS TAXCLASS4,
                    1 AS TAXCLASS5,
                    0 AS SWTAXINCL1,
                    0 AS SWTAXINCL2,
                    0 AS SWTAXINCL3,
                    0 AS SWTAXINCL4,
                    0 AS SWTAXINCL5,
                    1 AS RATETAX1,
                    1 AS RATETAX2,
                    1 AS RATETAX3,
                    1 AS RATETAX4,
                    1 AS RATETAX5,
                    0 AS AMTTAX1,
                    0 AS AMTTAX2,
                    0 AS AMTTAX3,
                    0 AS AMTTAX4,
                    0 AS AMTTAX5,
                    0 AS MISCTAX1,
                    0 AS MISCTAX2,
                    0 AS MISCTAX3,
                    0 AS MISCTAX4,
                    0 AS MISCTAX5,
                    0 AS GLTAXAMT1,
                    0 AS GLTAXAMT2,
                    0 AS GLTAXAMT3,
                    0 AS GLTAXAMT4,
                    0 AS GLTAXAMT5,
                    0 AS BKTAXAMT1,
                    0 AS BKTAXAMT2,
                    0 AS BKTAXAMT3,
                    0 AS BKTAXAMT4,
                    0 AS BKTAXAMT5,
                    0 AS TCAMTINCTX,
                    0 AS GLAMTINCTX,
                    0 AS BKAMTINCTX,
                    0 AS MSAMTINCTX,
                    NULL AS MISCAMOUNT,
                    NULL AS GLAMOUNT,
                    NULL AS BKAMOUNT,
                    0 AS TOTAPPLAMT,
                    0 AS TOTAPPLDSC,
                    0 AS TOTUNAPPL,
                    0 AS PJCCOST,
                    0 AS NOSUBDETL,
                    0 AS 'VALUES',
                    0 AS PROCESSCMD,
                    0 AS AMTTAXREC1,
                    0 AS AMTTAXREC2,
                    0 AS AMTTAXREC3,
                    0 AS AMTTAXREC4,
                    0 AS AMTTAXREC5,
                    0 AS AMTTAXEXP1,
                    0 AS AMTTAXEXP2,
                    0 AS AMTTAXEXP3,
                    0 AS AMTTAXEXP4,
                    0 AS AMTTAXEXP5,
                    0 AS AMTTAXTOBE,
                    0 AS TXAMT1RC,
                    0 AS TXAMT2RC,
                    0 AS TXAMT3RC,
                    0 AS TXAMT4RC,
                    0 AS TXAMT5RC,
                    0 AS TXTOTRC,
                    0 AS TXALLRC,
                    0 AS TXEXP1RC,
                    0 AS TXEXP2RC,
                    0 AS TXEXP3RC,
                    0 AS TXEXP4RC,
                    0 AS TXEXP5RC,
                    0 AS TXREC1RC,
                    0 AS TXREC2RC,
                    0 AS TXREC3RC,
                    0 AS TXREC4RC,
                    0 AS TXREC5RC,
                    0 AS TXBSE1HC,
                    0 AS TXBSE2HC,
                    0 AS TXBSE3HC,
                    0 AS TXBSE4HC,
                    0 AS TXBSE5HC,
                    0 AS TXREC1HC,
                    0 AS TXREC2HC,
                    0 AS TXREC3HC,
                    0 AS TXREC4HC,
                    0 AS TXREC5HC,
                    0 AS TXEXP1HC,
                    0 AS TXEXP2HC,
                    0 AS TXEXP3HC,
                    0 AS TXEXP4HC,
                    0 AS TXEXP5HC,
                    0 AS TXALLHC,
                    0 AS TXALL1HC,
                    0 AS TXALL2HC,
                    0 AS TXALL3HC,
                    0 AS TXALL4HC,
                    0 AS TXALL5HC,
                    0 AS TXALL1TC,
                    0 AS TXALL2TC,
                    0 AS TXALL3TC,
                    0 AS TXALL4TC,
                    0 AS TXALL5TC,
                    0 AS TXEXCLTC,
                    0 AS TXEXCLHC,
                    0 AS TXEXCLBC,
                    0 AS TXEXCLMC,
                    0 AS REVUNIQ,
                    0 AS NEWREVUNIQ,
                    NULL AS RVDETAILNO,
                    'FALSE' AS SWNODEFS,
                    241173908 AS CBBTHDRVH,
                    241173908 AS CBBTHDVW,
                    241693004 AS CBBTSDRVH,
                    241693004 AS CBBTSDVW,
                    0 AS CBERRNUM,
                    NULL AS CBERRDESC,
                    NULL AS CBERRFUNC,
                    'FALSE' AS SWNOTAXVAL,
                    'FALSE' AS SWREFUND,
                    'FALSE' AS SWZEROENT,
                    'FALSE' AS SWAUTALLOC,
                    'FALSE' AS SWINCLDISC,
                    'FALSE' AS SWIGNORE
               FROM ST_Register
               WHERE SalesDate = '2024-05-01 00:00:00.000') p
            UNPIVOT
               (DTLAMOUNT FOR ACCTDESC IN
                  (Deposit1, Deposit2, Deposit3, Deposit4, Deposit5, Deposit6, Deposit7, Deposit8)
            )AS unpvt;"""
    }

    # Load data from SQL Database into DataFrames
    data_frames = {}
    for table_name, query in queries.items():
        data_frames[table_name] = pd.read_sql(query, conn)

    # CSVs for the conversions
    csvs = {
        'Store_Detail': "Store_Detail.csv",
        'Account_Detail': "Account_Detail.csv"
    }

    # Load data from CSV files into separate DataFrames
    data_frames2 = {}
    for table_name, csv in csvs.items():
        data_frames2[table_name] = pd.read_csv(csv)

    # Assign Deposit Amount to multiple columns
    data_frames['Batch_Detail'].DEBITAMT = data_frames['Batch_Detail'].DTLAMOUNT
    data_frames['Batch_Detail'].MISCAMOUNT = data_frames['Batch_Detail'].DTLAMOUNT
    data_frames['Batch_Detail'].GLAMOUNT = data_frames['Batch_Detail'].DTLAMOUNT
    data_frames['Batch_Detail'].BKAMOUNT = data_frames['Batch_Detail'].DTLAMOUNT

    # Assign Batch ID to Header and Detail DataFrames
    batch_id = random.randint(0, 999999)
    data_frames['Batch_Header'].BATCHID = batch_id
    data_frames['Batch_Detail'].BATCHID = batch_id

    # Assign TEXTDESC on Batch Header sheet based on Store Detail DataFrame
    merged = pd.merge(data_frames['Batch_Header'], data_frames2['Store_Detail'], on='StoreID', how='left')
    merged['DATE1'] = pd.to_datetime(merged['DATE'])
    merged['DATE2'] = merged['DATE1'].dt.strftime('%d.%m.%Y')
    merged['TEXTDESC'] = merged['StoreKey'] + " BANK " + merged['DATE2']
    data_frames['Batch_Header'].TEXTDESC = merged['TEXTDESC']

    # Set ACCTDESC to CASH SHORTAGE or OVERAGE based on negative/positve value
    def cash_desc(ACCTDESC, DTLAMOUNT):
        if ACCTDESC == "CASH SHORTAGE/OVERAGE":
            if DTLAMOUNT < 0:
                return 'CASH SHORTAGE'
            else:
                return 'CASH OVERAGE'
        return ACCTDESC
    data_frames['Batch_Detail'].ACCTDESC = data_frames['Batch_Detail'].apply(lambda x: cash_desc(x['ACCTDESC'], x['DTLAMOUNT']), axis=1)

    # Assign column values on Batch Detail sheet based on Store and Account Detail DataFrames
    merged = pd.merge(data_frames['Batch_Detail'], data_frames2['Store_Detail'], on='StoreID', how='left')
    merged = pd.merge(merged, data_frames2['Account_Detail'], on='ACCTDESC', how='left')

    # TEXTDESC
    merged['DATE1'] = pd.to_datetime(merged['DATE'])
    merged['DATE2'] = merged['DATE1'].dt.strftime('%d.%m.%Y')
    merged['TEXTDESC'] = merged['StoreKey'] + " " + merged['AccountKey'] + " " + merged['DATE2']
    data_frames['Batch_Detail'].TEXTDESC = merged['TEXTDESC']

    # ACCTID
    data_frames['Batch_Detail'].ACCTID = merged['AcountID']

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
