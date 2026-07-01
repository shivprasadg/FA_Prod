SELECT
  Vendors.VendorSNm AS Vendor,
  nz([InvoiceNumber], "none") AS Invoice,
  Sum(vw_SixKeys.UnitCost) AS InvCost,
  Count(vw_SixKeys.AssetID) AS Units,
  CDbl(
    NZ([DrawnUnitInvoices].[Units], 0)
  ) AS Drawn,
  vw_SixKeys.SchID
FROM
  (
    vw_SixKeys
    LEFT JOIN Vendors ON vw_SixKeys.VendorId = Vendors.VendorID
  )
  LEFT JOIN DrawnUnitInvoices ON (
    vw_SixKeys.VendorId = DrawnUnitInvoices.VendorID
  )
  AND (
    vw_SixKeys.InvoiceNumber = DrawnUnitInvoices.Invoice
  )
GROUP BY
  Vendors.VendorSNm,
  nz([InvoiceNumber], "none"),
  CDbl(
    NZ([DrawnUnitInvoices].[Units], 0)
  ),
  vw_SixKeys.SchID
ORDER BY
  nz([InvoiceNumber], "none");
