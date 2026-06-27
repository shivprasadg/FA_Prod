SELECT
  DISTINCT First(ChildParts.ChildVendorREF) AS DecalVendor,
  ChildParts.UnitREF AS AssetId,
  ChildParts.childFinalCost AS Decal,
  ChildParts.ChildInvoiceNum AS DecalInvoice,
  ChildParts.ChildVendorPytDate AS DecalPaid
FROM
  ScheduleUnitDocs3
  INNER JOIN ChildParts ON ScheduleUnitDocs3.AssetId = ChildParts.UnitREF
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Decal*"
    )
  )
GROUP BY
  ChildParts.UnitREF,
  ChildParts.childFinalCost,
  ChildParts.ChildInvoiceNum,
  ChildParts.ChildVendorPytDate
ORDER BY
  First(ChildParts.ChildVendorREF);
