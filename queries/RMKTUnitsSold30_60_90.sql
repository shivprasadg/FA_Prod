SELECT
  Rmkt.RmkSm,
  Count(Rmkt.UnitRef) AS UnitsSold,
  Sum(
    IIf(
      [InvoicePaidDate] Between Date()
      And Date()-31,
      1,
      0
    )
  ) AS UnitsSoldLast30,
  Sum(
    IIf(
      [InvoicePaidDate] Between Date()-31
      And Date()-62,
      1,
      0
    )
  ) AS UnitsSoldLast60,
  Sum(
    IIf(
      [InvoicePaidDate] Between Date()-63
      And Date()-93,
      1,
      0
    )
  ) AS UnitsSoldLast90,
  Sum(
    IIf(
      [InvoicePaidDate] Between Date()
      And Date()-31,
      [SoldPrice],
      0
    )
  ) AS AmtSold30,
  Sum(
    IIf(
      [InvoicePaidDate] Between Date()-31
      And Date()-62,
      [SoldPrice],
      0
    )
  ) AS AmtSold60,
  Sum(
    IIf(
      [InvoicePaidDate] Between Date()-63
      And Date()-93,
      [SoldPrice],
      0
    )
  ) AS AmtSold90
FROM
  Rmkt
WHERE
  (
    (
      (Rmkt.InvoicePaidDate) Is Not Null
    )
  )
GROUP BY
  Rmkt.RmkSm;
