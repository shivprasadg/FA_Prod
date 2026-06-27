SELECT
  DISTINCT ChildParts.ChildVendorREF AS DecalVendor,
  Vendors.VendorSNm AS Name,
  ScheduleUnitDocs3.AssetId,
  ChildParts.childFinalCost AS CoNameCost,
  ChildParts.ChildInvoiceNum AS CoNameInvoice,
  ChildParts.ChildVendorPytDate AS CoNamePaid,
  ChildParts.ChildClass AS DecalType
FROM
  (
    ScheduleUnitDocs3
    INNER JOIN ChildParts ON ScheduleUnitDocs3.AssetId = ChildParts.UnitREF
  )
  INNER JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Decal*"
    )
  )
GROUP BY
  ChildParts.ChildVendorREF,
  Vendors.VendorSNm,
  ScheduleUnitDocs3.AssetId,
  ChildParts.childFinalCost,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate,
  ChildParts.ChildClass
HAVING
  (
    (
      (ChildParts.ChildClass)= "Co. Name"
    )
  )
ORDER BY
  ChildParts.ChildVendorREF;
