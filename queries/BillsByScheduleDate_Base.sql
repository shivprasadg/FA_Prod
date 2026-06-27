SELECT
  Bills.BillSchRef AS BilledSchID,
  Bills.BillType,
  Min(Bills.Billfrom) AS BilledFrom,
  Max(Bills.Billto) AS BilledTo
FROM
  Bills
GROUP BY
  Bills.BillSchRef,
  Bills.BillType;
