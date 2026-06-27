SELECT
  DISTINCT First(ChildParts.ChildVendorREF) AS DecalVendor,
  Max(ChildParts.ChildClass) AS DecalType,
  ScheduleUnitDocs3.AssetId INTO Cost_DecalVendor3
FROM
  (
    (
      (
        ScheduleUnitDocs3
        INNER JOIN ChildParts ON ScheduleUnitDocs3.AssetId = ChildParts.UnitREF
      )
      INNER JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID
    )
    LEFT JOIN Cost_DecalVendor1 ON ChildParts.ChildClass = Cost_DecalVendor1.DecalType
  )
  LEFT JOIN Cost_DecalVendor2 ON ChildParts.ChildClass = Cost_DecalVendor2.DecalType
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Decal*"
    )
    AND (
      (Cost_DecalVendor1.DecalType) Is Null
    )
    AND (
      (Cost_DecalVendor2.DecalType) Is Null
    )
  )
GROUP BY
  ScheduleUnitDocs3.AssetId
ORDER BY
  First(ChildParts.ChildVendorREF);
