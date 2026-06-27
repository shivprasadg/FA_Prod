SELECT
  RemarketExpense.[AssetId],
  RemarketExpense.[VendorId],
  RemarketExpense.[InvoiceNumber],
  RemarketExpense.[ExpenseTypeId],
  RemarketExpense.[RMKTExpenseId],
  RemarketExpense.[Expense],
  RemarketExpense.[ExpenseDate]
FROM
  RemarketExpense
WHERE
  (
    (
      (RemarketExpense.[AssetId]) In (
        SELECT
          [AssetId]
        FROM
          [RemarketExpense] As Tmp
        GROUP BY
          [AssetId],
          [VendorId],
          [InvoiceNumber],
          [ExpenseTypeId]
        HAVING
          Count(*)> 1
          And [VendorId] = [RemarketExpense].[VendorId]
          And [InvoiceNumber] = [RemarketExpense].[InvoiceNumber]
          And [ExpenseTypeId] = [RemarketExpense].[ExpenseTypeId]
      )
    )
  )
ORDER BY
  RemarketExpense.[AssetId],
  RemarketExpense.[VendorId],
  RemarketExpense.[InvoiceNumber],
  RemarketExpense.[ExpenseTypeId];
