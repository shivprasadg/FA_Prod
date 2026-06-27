SELECT
  SchGrp.SGrpSchID AS SchID,
  Units.UnitID,
  Vendors.VendorID AS ChildVendorID,
  Vendors.VendorName AS ChildVendor,
  ChildParts.childFinalCost AS ChildCost,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildType
FROM
  (
    (
      Units
      LEFT JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
    )
    LEFT JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
  )
  LEFT JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID;
