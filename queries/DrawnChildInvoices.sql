SELECT
  Vendors.vendorsnm AS Vendor,
  nz(DCP.ChildInvoiceNum, "none") AS Invoice,
  Sum(DCP.childCost) AS TCost,
  Count(DCP.ChildPartId) AS Parts,
  DCP.ChildDrawDownId,
  DCP.ChildVendorID
FROM
  vw_Sch_Grp_Units_ChildParts AS DCP
  INNER JOIN Vendors ON DCP.ChildVendorID = Vendors.VendorID
GROUP BY
  Vendors.vendorsnm,
  nz(DCP.ChildInvoiceNum, "none"),
  DCP.ChildDrawDownId,
  DCP.ChildVendorID
HAVING
  (
    (
      (DCP.ChildDrawDownId) Is Not Null
    )
  );
