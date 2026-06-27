SELECT
  RemarketExpense.AssetId,
  RemarketExpense.VendorId,
  RemarketExpense.InvoiceNumber,
  RemarketExpense.ExpenseTypeId,
  RemarketExpense.RMKTExpenseId,
  RemarketExpense.Expense,
  RemarketExpense.ExpenseDate,
  Vendors.VendorName
FROM
  RemarketExpense
  INNER JOIN Vendors ON RemarketExpense.VendorId = Vendors.VendorID
WHERE
  (
    (
      (RemarketExpense.AssetId) In (
        SELECT
          [AssetId]
        FROM
          [RemarketExpense] As Tmp
        GROUP BY
          [AssetId],
          [InvoiceNumber],
          [ExpenseTypeId]
        HAVING
          Count(*)> 1
          And [InvoiceNumber] = [RemarketExpense].[InvoiceNumber]
          And [ExpenseTypeId] = [RemarketExpense].[ExpenseTypeId]
      )
    )
  )
ORDER BY
  RemarketExpense.AssetId,
  RemarketExpense.VendorId,
  RemarketExpense.InvoiceNumber,
  RemarketExpense.ExpenseTypeId;
