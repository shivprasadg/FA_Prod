SELECT
  DISTINCT Max(ChildParts.ChildVendorREF) AS DecalVendor,
  ScheduleUnitDocs3.AssetId INTO Cost_DecalVendor4
FROM
  (
    (
      (
        (
          ScheduleUnitDocs3
          INNER JOIN ChildParts ON ScheduleUnitDocs3.AssetId = ChildParts.UnitREF
        )
        INNER JOIN Vendors ON ChildParts.ChildVendorREF = Vendors.VendorID
      )
      LEFT JOIN Cost_DecalVendor1 ON ChildParts.ChildVendorREF = Cost_DecalVendor1.DecalVendor
    )
    LEFT JOIN Cost_DecalVendor2 ON ChildParts.ChildVendorREF = Cost_DecalVendor2.DecalVendor
  )
  LEFT JOIN Cost_DecalVendor3 ON ChildParts.ChildVendorREF = Cost_DecalVendor3.DecalVendor
WHERE
  (
    (
      (ChildParts.ChildType) Like "*Decal*"
    )
    AND (
      (Cost_DecalVendor1.DecalVendor) Is Null
    )
    AND (
      (Cost_DecalVendor2.DecalVendor) Is Null
    )
    AND (
      (Cost_DecalVendor3.DecalVendor) Is Null
    )
  )
GROUP BY
  ScheduleUnitDocs3.AssetId
ORDER BY
  Max(ChildParts.ChildVendorREF);
