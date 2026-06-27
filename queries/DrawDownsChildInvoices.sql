SELECT
  DrawDownsChildInvoices0.Invoice,
  DrawDownsChildInvoices0.Vendor,
  DrawDownsChildInvoices0.InvCost,
  [Parts] - CDbl(
    Nz([Drawn], 0)
  ) AS Avail,
  CDbl(
    Nz(
      [DrawDownsChildInvoices0].[Drawn],
      0
    )
  ) AS Drawn,
  DrawDownsChildInvoices0.Parts,
  DrawDownsChildInvoices0.SchID,
  DrawDownsChildInvoices0.SGrpID AS GroupID
FROM
  DrawDownsChildInvoices0
GROUP BY
  DrawDownsChildInvoices0.Invoice,
  DrawDownsChildInvoices0.Vendor,
  DrawDownsChildInvoices0.InvCost,
  [Parts] - CDbl(
    Nz([Drawn], 0)
  ),
  CDbl(
    Nz(
      [DrawDownsChildInvoices0].[Drawn],
      0
    )
  ),
  DrawDownsChildInvoices0.Parts,
  DrawDownsChildInvoices0.SchID,
  DrawDownsChildInvoices0.SGrpID
ORDER BY
  DrawDownsChildInvoices0.Invoice;
