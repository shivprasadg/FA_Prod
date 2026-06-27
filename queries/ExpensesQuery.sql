SELECT
  Expenses.ExpenseDate,
  Expenses.ExpenseTypeRef,
  Expenses.ExpenseID,
  Expenses.DealsRef,
  Expenses.GroupsRef,
  Expenses.ExpenseRep,
  Expenses.ExpenseDesc,
  Expenses.ExpenseMemo,
  Expenses.ExpenseAmt
FROM
  Expenses
ORDER BY
  Expenses.ExpenseDate DESC,
  Expenses.ExpenseTypeRef;
