SELECT
  DrawableUnitInvoices0.Vendor AS Expr1,
  DrawableUnitInvoices0.Invoice AS Expr2,
  DrawableUnitInvoices0.InvCost AS Expr3,
  [Units] - CDbl(
    Nz([Drawn], 0)
  ) AS Avail,
  CDbl(
    Nz(
      [DrawableUnitInvoices0].[Drawn],
      0
    )
  ) AS Drawn,
  DrawableUnitInvoices0.Units AS Expr4,
  DrawableUnitInvoices0.SchID AS Expr5
FROM
  DrawableUnitInvoices0
GROUP BY
  DrawableUnitInvoices0.Vendor,
  DrawableUnitInvoices0.Invoice,
  DrawableUnitInvoices0.InvCost,
  [Units] - CDbl(
    Nz([Drawn], 0)
  ),
  CDbl(
    Nz(
      [DrawableUnitInvoices0].[Drawn],
      0
    )
  ),
  DrawableUnitInvoices0.Units,
  DrawableUnitInvoices0.SchID
ORDER BY
  DrawableUnitInvoices0.Invoice;
