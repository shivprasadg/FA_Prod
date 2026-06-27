SELECT
  DISTINCT ChildParts.UnitREF AS AssetId,
  ChildParts.ChildVendorREF AS DecalVendor,
  Vendors.VendorSNm AS Name,
  ChildParts.childFinalCost AS RegDecalCost,
  ChildParts.ChildInvoiceNum AS RegDecalInvoice,
  ChildParts.ChildVendorPytDate AS RegDecalPaid,
  ChildParts.ChildClass AS DecalType,
  ChildParts.DrawDownId AS RegDecalDraw
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
      (ChildParts.ChildClass) Like "*Reg*"
    )
  )
ORDER BY
  ChildParts.ChildVendorREF;
