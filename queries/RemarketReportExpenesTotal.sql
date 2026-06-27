SELECT
  RemarketExpense.AssetId,
  Sum(RemarketExpense.Expense) AS SubTotalExpense
FROM
  RemarketExpense
GROUP BY
  RemarketExpense.AssetId;
