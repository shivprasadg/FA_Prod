SELECT
  DrawDowns.DrawDownId,
  Vendors.VendorSNm AS Vendor,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  DrawDowns.DrawRequestNumber AS RequestNumber,
  DrawDowns.DrawName,
  DrawDowns.DrawNumber AS Draw,
  Units.UnitID AS AssetID,
  Units.UnitVIN AS VIN,
  Units.UnitUnitNum AS UnitNum,
  [UnitYr] & " " & [UnitMake] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& IIf(
    [UnitType] = "Tractor", "", " " & [UnitType]
  ) AS UnitDesc,
  Units.UnitFinalCost AS UnitCost,
  Units.UnitInvoiceNum AS Invoice,
  Units.UnitInvoiceDate AS InvoiceDate,
  vw_SixKeys.UnitDescDocVer,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  vw_SixKeys.VendorId,
  DrawDowns.DrawRequestNumber
FROM
  (
    DrawDowns
    INNER JOIN (
      Units
      INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
    ) ON DrawDowns.DrawDownId = vw_SixKeys.DrawDownId
  )
  INNER JOIN Vendors ON vw_SixKeys.VendorId = Vendors.VendorID
WHERE
  (
    (
      (DrawDowns.DrawDownId) Is Not Null
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  DrawDowns.DrawRequestNumber,
  DrawDowns.DrawNumber,
  Units.UnitInvoiceNum,
  Units.UnitId;
