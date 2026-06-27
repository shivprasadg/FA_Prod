SELECT
  SchGrp.SGrpSchID AS SchID,
  Units.UnitID,
  Vendors.VendorID,
  Vendors.VendorName AS UnitVendor,
  Units.Unitfinalcost AS UnitCost,
  Units.UnitInvoiceNum,
  Units.UnitType
FROM
  (
    Units
    LEFT JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
  )
  INNER JOIN Vendors ON Units.VendorREF = Vendors.VendorID;
