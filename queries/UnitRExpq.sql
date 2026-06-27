SELECT
  Units.UnitID,
  Vendors.VendorName,
  RemarketExpenseType.RemarketExpense,
  RmkExpense.InvAmt,
  RmkExpense.InvoiceNumber,
  RmkExpense.InvoiceDate
FROM
  (
    (
      RmkExpense
      RIGHT JOIN (
        Units
        INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
      ) ON RmkExpense.RmktRef = Rmkt.RmktId
    )
    LEFT JOIN Vendors ON RmkExpense.VendorRef = Vendors.VendorID
  )
  LEFT JOIN RemarketExpenseType ON RmkExpense.ExpenseTypeId = RemarketExpenseType.RemarketExpenseTypeId;
