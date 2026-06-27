SELECT
  vw_SixKeys.Schedule,
  vw_SixKeys.Unitnum,
  vw_SixKeys.VIN,
  vw_SixKeys.UnitDescDocVer,
  [VendorName] & IIf(
    [UnitType] Like "*tax*", " (" & [UnitType] & ")",
    ""
  ) AS Vendor,
  SchDocsVendorUnion.InvoiceNum,
  SchDocsVendorUnion.UnitType,
  SchDocsVendorUnion.cost AS Cost,
  vw_SixKeys.VendorCost
FROM
  vw_SixKeys
  INNER JOIN SchDocsVendorUnion ON vw_SixKeys.AssetID = SchDocsVendorUnion.UnitID
ORDER BY
  [VendorName] & IIf(
    [UnitType] Like "*tax*", " (" & [UnitType] & ")",
    ""
  ),
  SchDocsVendorUnion.cost;
