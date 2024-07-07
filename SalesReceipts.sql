SELECT
    NULL AS 'Sales Receipt No',
    'Cash Sales' AS 'Customer',
    CONVERT(varchar,MAX(r.SalesDate),103) AS 'Sales Receipt Date',
    10300 AS 'Deposit To',
    'Cash' AS 'Payment method',
    'NotApplicable' AS 'Global Tax Calculation',
    NULL AS 'Product/Service',
    (SUM(r.Sales1) + SUM(r.Sales2) + SUM(r.Sales7)) 'Cash Sales - Other',
    SUM(r.Sales6) AS 'Cash Sales - Beverages',
    (SUM(r.Sales3) + SUM(r.Sales6)) AS 'Cash Sales - Juici',
    (SUM(r.Tax1)) AS '12260 GCT Payable/Receivable',
    NULL AS 'Product/Service Description',
    NULL AS 'Product/Service Amount',
    'JUICI' AS 'Product/Service Class'
FROM ST_Register AS r
WHERE (r.SalesDate BETWEEN DATEADD(DAY,-7,'INVOICE_DATE')
    AND 'INVOICE_DATE')
GROUP BY SiteNumber