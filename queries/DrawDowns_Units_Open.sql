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
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  vw_SixKeys.UnitNum,
  vw_SixKeys.UnitDescShortVer AS UnitDesc,
  vw_SixKeys.UnitCost,
  vw_SixKeys.OECUnit AS OEC,
  vw_SixKeys.UnitInvoiceNum AS Invoice,
  vw_SixKeys.UnitinvoiceDate AS InvoiceDate,
  vw_SixKeys.UnitDescDocVer,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  vw_SixKeys.VendorId,
  vw_SixKeys.MLOrig
FROM
  (
    DrawDowns
    RIGHT JOIN vw_SixKeys ON DrawDowns.DrawDownId = vw_SixKeys.DrawDownId
  )
  INNER JOIN Vendors ON vw_SixKeys.VendorId = Vendors.VendorID
WHERE
  (
    (
      (DrawDowns.DrawDownId) Is Null
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName,
  DrawDowns.DrawRequestNumber,
  DrawDowns.DrawNumber,
  vw_SixKeys.UnitInvoiceNum,
  vw_SixKeys.UnitId;
