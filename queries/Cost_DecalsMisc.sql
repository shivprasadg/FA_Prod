SELECT
  DISTINCT ChildParts.UnitREF AS AssetId,
  ChildParts.ChildVendorREF AS DecalVendor,
  Vendors.VendorSNm AS Name,
  ChildParts.childFinalCost AS MiscCost,
  ChildParts.ChildInvoiceNum AS MiscInvoice,
  ChildParts.ChildVendorPytDate AS MiscPaid,
  ChildParts.ChildClass AS DecalType,
  ChildParts.DrawDownId AS MiscDraw
FROM
  ChildParts
  INNER JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Decal*"
    )
  )
GROUP BY
  ChildParts.UnitREF,
  ChildParts.ChildVendorREF,
  Vendors.VendorSNm,
  ChildParts.childFinalCost,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.ChildClass,
  ChildParts.DrawDownId
HAVING
  (
    (
      (ChildParts.ChildClass)= "Misc."
    )
  )
ORDER BY
  ChildParts.ChildVendorREF;
