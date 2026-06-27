SELECT
  Bills.BillSchRef,
  Max(Bills.BillSent) AS MaxOfBillSent
FROM
  Bills
GROUP BY
  Bills.BillSchRef,
  Bills.BillType
HAVING
  (
    (
      (Bills.BillType)= "PerDiem"
    )
  );
