SELECT
  RemarketExpense.AssetId,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 1, [Expense], 0)
    )
  ) AS Title,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 2, [Expense], 0)
    )
  ) AS Transport,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 4, [Expense], 0)
    )
  ) AS Advertising,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 5, [Expense], 0)
    )
  ) AS Repairs,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 6, [Expense], 0)
    )
  ) AS Paint,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 7, [Expense], 0)
    )
  ) AS Bankshare,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 8, [Expense], 0)
    )
  ) AS Interest,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 9, [Expense], 0)
    )
  ) AS Commission,
  Sum(
    CCur(
      IIf([ExpenseTypeId] = 10, [Expense], 0)
    )
  ) AS Other,
  [Title] + [Transport] + [Advertising] + [Repairs] + [Paint] + [Bankshare] + [Interest] + [Commission] + [Other] AS SubTotalExpense
FROM
  RemarketExpense
GROUP BY
  RemarketExpense.AssetId;
