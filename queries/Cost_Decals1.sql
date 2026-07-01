SELECT
  Cost_DecalVendor1.AssetId,
  Vendors.VendorName AS DecalVendor1,
  ChildParts.childFinalCost AS Decal1Cost,
  ChildParts.ChildInvoiceNum AS Decal1Invoice,
  ChildParts.ChildVendorPytDate AS Decal1Paid
FROM
  (
    Cost_DecalVendor1
    INNER JOIN ChildParts ON (
      Cost_DecalVendor1.AssetId = ChildParts.UnitREF
    )
    AND (
      Cost_DecalVendor1.DecalVendor = ChildParts.ChildVendorREF
    )
  )
  INNER JOIN Vendors ON Cost_DecalVendor1.DecalVendor = Vendors.VendorID;
