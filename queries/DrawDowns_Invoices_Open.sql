SELECT
  Units.DrawDownId,
  Vendors.vendorsnm AS Vendor,
  Nz(
    [Units.UnitInvoiceNum], "Pending"
  ) AS Invoice,
  Count(Units.UnitId) AS Units,
  Sum(Units.VendorCost) AS TCost,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.VendorID,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID
FROM
  (
    Units
    INNER JOIN Vendors ON Units.VendorREF = Vendors.VendorID
  )
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
GROUP BY
  Units.DrawDownId,
  Vendors.vendorsnm,
  Nz(
    [Units.UnitInvoiceNum], "Pending"
  ),
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.VendorID,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID
HAVING
  (
    (
      (Units.DrawDownId) Is Null
    )
  );
