SELECT
    NULL AS 'Sales Receipt No',
    'Cash Sales' AS 'Customer',
    r.SalesDate AS 'Sales Receipt Date',
    10300 AS 'Deposit To',
    'Cash' AS 'Payment method',
    'NotApplicable' AS 'Global Tax Calculation',
    NULL AS 'Product/Service',
    (r.Sales1 + r.Sales2 + r.Sales7) 'Cash Sales - Other',
    (r.Sales6) AS 'Cash Sales - Beverages',
    (r.Sales3 + r.Sales6) AS 'Cash Sales - Juici',
    (r.Tax1) AS '12260 GCT Payable/Receivable',
    NULL AS 'Product/Service Description',
    NULL AS 'Product/Service Amount',
    'JUICI' AS 'Product/Service Class'
FROM ST_Register AS r
WHERE (r.SalesDate BETWEEN DATEADD(DAY,-7,'INVOICE_DATE')
    AND 'INVOICE_DATE')