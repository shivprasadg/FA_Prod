SELECT
  Units.DrawDownId,
  Vendors.vendorsnm AS Vendor,
  Nz([UnitInvoiceNum], "none") AS Invoice,
  Count(Units.UnitId) AS Units,
  Sum(Units.UnitFinalCost) AS TCost,
  Vendors.VendorID
FROM
  Units
  INNER JOIN Vendors ON Units.VendorREF = Vendors.VendorID
GROUP BY
  Units.DrawDownId,
  Vendors.vendorsnm,
  Nz([UnitInvoiceNum], "none"),
  Vendors.VendorID
HAVING
  (
    (
      (Units.DrawDownId) Is Not Null
    )
  );
