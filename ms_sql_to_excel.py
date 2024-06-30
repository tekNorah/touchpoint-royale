import random
import pyodbc
import pandas as pd
import os
import argparse
from datetime import datetime

def main(start_date, end_date):
    # Parse the input dates
    try:
        start_date = datetime.strptime(start_date, "%Y-%m-%d")
        end_date = datetime.strptime(end_date, "%Y-%m-%d")
    except ValueError as e:
        print(f"Error parsing dates: {e}")
        return

    # Print the dates (for debugging purposes)
    print(f"Start Date: {start_date}")
    print(f"End Date: {end_date}")

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
                    0 AS NODETAILS,
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
                WHERE SalesDate >= '""" + start_date.strftime("%Y-%m-%d %H:%M:%S") + """'
                  AND SalesDate <= '""" + end_date.strftime("%Y-%m-%d %H:%M:%S") + """'
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
                    ACCTQTYSW,
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
                    BKAMOUNT
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
                        0 AS DEBITAMT,
                        0 AS CREDITAMT,
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
                        NULL AS BKAMOUNT
                   FROM ST_Register
                   WHERE SalesDate >= '""" + start_date.strftime("%Y-%m-%d %H:%M:%S") + """'
                     AND SalesDate <= '""" + end_date.strftime("%Y-%m-%d %H:%M:%S") + """') p
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
            'Account_Detail': "Account_Detail.csv",
            'Batch_Miscellaneous': 'Batch_Miscellaneous.csv',
            'Batch_Sub_Detail': 'Batch_Sub_Detail.csv',
            'Batch_Detail_Adjustments': 'Batch_Detail_Adjustments.csv',
            'Batch_Header_Optional_Fields': 'Batch_Header_Optional_Fields.csv',
            'Batch_Detail_Optional_Fields': 'Batch_Detail_Optional_Fields.csv'
        }

        # Load data from CSV files into separate DataFrames
        data_frames2 = {}
        for table_name, csv in csvs.items():
            data_frames2[table_name] = pd.read_csv(csv)

        # Assign Deposit Amount to multiple columns
        def debit_amount(DTLAMOUNT):
            if DTLAMOUNT > 0:
                return DTLAMOUNT
            return 0
        def credit_amount(DTLAMOUNT):
            if DTLAMOUNT < 0:
                return DTLAMOUNT
            return 0
        data_frames['Batch_Detail'].DEBITAMT = data_frames['Batch_Detail'].apply(lambda x: debit_amount(x['DTLAMOUNT']), axis=1)
        data_frames['Batch_Detail'].CREDITAMT = data_frames['Batch_Detail'].apply(lambda x: credit_amount(x['DTLAMOUNT']), axis=1)
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

        # Append Store Key to Account ID when post fix is missing
        data_frames['Batch_Detail'].ACCTID = data_frames['Batch_Detail'].apply(
            lambda row: row.ACCTID + "-" + row.TEXTDESC[:4] if len(row.ACCTID) == 4 else row.ACCTID,
            axis=1
        )

        # Assign NODETAILS based on count of matching Batch Detail records on ENTRYNO
        def count_details(header_entryno):
          return data_frames['Batch_Detail'][data_frames['Batch_Detail'].ENTRYNO == header_entryno].shape[0]

        data_frames['Batch_Header'].NODETAILS = data_frames['Batch_Header'].apply(lambda x: count_details(x['ENTRYNO']), axis=1)

        # Assign HDRDEBIT and HDRCREDIT on Batch Header Datafram based on sum of DEBITAMT and CREDITAMT on Batch Detail
        def sum_debit(entryno):
            return data_frames['Batch_Detail'].loc[data_frames['Batch_Detail'].ENTRYNO == entryno, 'DEBITAMT'].sum()
        def sum_credit(entryno):
            return data_frames['Batch_Detail'].loc[data_frames['Batch_Detail'].ENTRYNO == entryno, 'CREDITAMT'].sum()

        data_frames['Batch_Header'].HDRDEBIT = data_frames['Batch_Header'].apply(lambda x: sum_debit(x['ENTRYNO']), axis=1)
        data_frames['Batch_Header'].HDRCREDIT = data_frames['Batch_Header'].apply(lambda x: sum_credit(x['ENTRYNO']), axis=1)

        # Create Batch_Miscellaneous DataFrame as subset of Batch Header DataFrame
        data_frames['Batch_Miscellaneous'] = data_frames['Batch_Header'][["BATCHID", "ENTRYNO"]]

        # Add columns from Batch_Miscellaneous DataFrame2 to Batch_Miscellaneous DataFrame
        data_frames['Batch_Miscellaneous'] = data_frames['Batch_Miscellaneous'].join(data_frames2['Batch_Miscellaneous'])

        # Set Defaults for some columns on Batch_Miscellaneous DataFrame
        data_frames['Batch_Miscellaneous'].DETAILNO = '0000000200'
        data_frames['Batch_Miscellaneous'].SWKEEPTOT = 'FALSE'
        data_frames['Batch_Miscellaneous'].ACCTROW = 1
        data_frames['Batch_Miscellaneous'].SWAPPROVED = 'FALSE'
        data_frames['Batch_Miscellaneous'].ACCTYPE = 0
        data_frames['Batch_Miscellaneous'].COVERTYPE = 0
        data_frames['Batch_Miscellaneous'].EITYPE = 0

        # Create remaining DataFrames based on CSV Templates
        data_frames['Batch_Sub_Detail'] = data_frames2['Batch_Sub_Detail']
        data_frames['Batch_Detail_Adjustments'] = data_frames2['Batch_Detail_Adjustments']
        data_frames['Batch_Header_Optional_Fields'] = data_frames2['Batch_Header_Optional_Fields']
        data_frames['Batch_Detail_Optional_Fields'] = data_frames2['Batch_Detail_Optional_Fields']

        # Remove fields not needed from data frames
        data_frames['Batch_Header'] = data_frames['Batch_Header'].drop(columns=['StoreID'])
        data_frames['Batch_Detail'] = data_frames['Batch_Detail'].drop(columns=['StoreID'])
        data_frames['Batch_Detail'] = data_frames['Batch_Detail'].drop(columns=['DATE'])

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

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="A script that takes two date parameters.")
    parser.add_argument('start_date', type=str, help='The start date in YYYY-MM-DD format')
    parser.add_argument('end_date', type=str, help='The end date in YYYY-MM-DD format')
    args = parser.parse_args()
    main(args.start_date, args.end_date)