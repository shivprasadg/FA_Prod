SELECT
  Units.UnitID,
  Vendors.VendorName,
  Vendors.VendorSNm,
  ChildParts.ChildVIN,
  ChildParts.ChildMake,
  ChildParts.ChildClass,
  ChildParts.ChildType,
  ChildParts.childFinalCost,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.ChildVendorREF
FROM
  (
    Units
    INNER JOIN ChildParts ON Units.UnitID = ChildParts.UnitREF
  )
  INNER JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID
WHERE
  (
    (
      (ChildParts.ChildVendorREF)= [vendorref]
    )
  );
