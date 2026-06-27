SELECT
  vw_SixKeys.Schedule,
  [VendorName] & IIf(
    [UnitType] Like "*tax*", " (" & [UnitType] & ")",
    ""
  ) AS Vendor,
  Sum(SchDocsVendorUnion.cost) AS Cost,
  Sum(vw_SixKeys.VendorCost) AS VendorCost
FROM
  vw_SixKeys
  INNER JOIN SchDocsVendorUnion ON vw_SixKeys.AssetID = SchDocsVendorUnion.UnitID
GROUP BY
  vw_SixKeys.Schedule,
  [VendorName] & IIf(
    [UnitType] Like "*tax*", " (" & [UnitType] & ")",
    ""
  )
ORDER BY
  [VendorName] & IIf(
    [UnitType] Like "*tax*", " (" & [UnitType] & ")",
    ""
  ),
  Sum(SchDocsVendorUnion.cost);
