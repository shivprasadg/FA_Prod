SELECT
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup AS [Group],
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  vw_SixKeys.MYear,
  vw_SixKeys.Make,
  vw_SixKeys.Model,
  vw_SixKeys.Type,
  vw_SixKeys.SubType,
  Units.UnitStatus,
  Units.UnitFinalCost,
  vw_SixKeys.PartsCost AS ChildPartCost,
  vw_SixKeys.OECUnit,
  vw_SixKeys.UnitInvoiceDate AS InvDate,
  vw_SixKeys.UnitInvoiceNum AS Invoice,
  vw_SixKeys.VendorCost,
  Units.UnitVendorPytDate AS VendorPaid,
  [Units].[VendorCost] + [ChildPartCost] AS VendorTotalCost,
  Vendors.VendorID,
  Vendors.VendorName
FROM
  Vendors
  INNER JOIN (
    Units
    INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
  ) ON Vendors.VendorID = vw_SixKeys.VendorId
WHERE
  (
    (
      (Vendors.VendorID)= [Forms]![VendorsForm]![VendorID]
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.UnitInvoiceDate DESC,
  Units.UnitVendorPytDate DESC;
