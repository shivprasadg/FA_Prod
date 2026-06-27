SELECT
  Vendors.vendorsnm AS Vendor,
  Nz(DCP.ChildInvoiceNum, "none") AS Invoice,
  Sum(DCP.childCost) AS InvCost,
  Count(DCP.ChildPartID) AS Parts,
  CDbl(
    Nz([DrawnChildInvoices].[Parts], 0)
  ) AS Drawn,
  DCP.SchID,
  DCP.SGrpID
FROM
  (
    vw_Sch_Grp_Units_ChildParts AS DCP
    LEFT JOIN Vendors ON DCP.ChildVendorID = Vendors.VendorID
  )
  LEFT JOIN DrawnChildInvoices ON (
    DCP.ChildVendorID = DrawnChildInvoices.ChildVendorID
  )
  AND (
    DCP.ChildInvoiceNum = DrawnChildInvoices.Invoice
  )
GROUP BY
  Vendors.vendorsnm,
  Nz(DCP.ChildInvoiceNum, "none"),
  CDbl(
    Nz([DrawnChildInvoices].[Parts], 0)
  ),
  DCP.SchID,
  DCP.SGrpID;
