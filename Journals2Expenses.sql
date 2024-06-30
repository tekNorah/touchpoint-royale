SELECT
    ih.InvoiceID AS 'Ref No',
    ih.InvoiceDate AS 'Payment Date',
    ih.Notes AS 'Memo',
    v.AccountNumber AS 'Account',
    ic.ItemCategoryNumber,
    ic.ItemCategoryName,
    NULL AS 'Expense Account',
    id.ExtendedPrice AS 'Expense Line Amount',
    id.RawMaterialItemCatName AS 'Expense Description',
    v.VendorName AS 'Payee',
    'JUICI' AS 'Expense Class',
    'Cash' AS 'Payment Method'
FROM ST_InvoiceHeaders AS ih
         INNER JOIN ST_Vendors AS v
                    ON (ih.VendorID = v.VendorID)
         INNER JOIN ST_InvoiceDetails id
                    ON (ih.InvoiceNumber = id.InvoiceNumber)
         INNER JOIN ST_ItemCategories AS ic
                    ON (id.RawMaterialItemCatNum = ic.ItemCategoryNumber)
WHERE v.VendorName IN ('Petty Cash',
                       'NCB MC 7604',
                       'Business Edge 2490',
                       'Business Elite 5240')
  AND ih.InvoiceDate = 'INVOICE_DATE'