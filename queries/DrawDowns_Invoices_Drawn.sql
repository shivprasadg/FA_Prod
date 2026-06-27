SELECT
  Units.DrawDownId,
  Vendors.vendorsnm AS Vendor,
  Units.UnitInvoiceNum AS Invoice,
  Count(Units.UnitId) AS Units,
  Sum(Units.UnitFinalCost) AS TCost,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.VendorID,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  DrawDowns.DrawRequestNumber,
  DrawDowns.DrawName,
  DrawDowns.DrawNumber,
  Sum(Units.unitIfamt) AS IFAmount,
  Banks.BankshortName AS IFBank
FROM
  (
    (
      (
        vw_SixKeys
        INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
      )
      LEFT JOIN Banks ON Units.UnitIFBank = Banks.BankShortName
    )
    INNER JOIN Vendors ON Units.VendorREF = Vendors.VendorID
  )
  LEFT JOIN DrawDowns ON Units.DrawDownId = DrawDowns.DrawDownId
GROUP BY
  Units.DrawDownId,
  Vendors.vendorsnm,
  Units.UnitInvoiceNum,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.VendorID,
  vw_SixKeys.clientID,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID,
  DrawDowns.DrawRequestNumber,
  DrawDowns.DrawName,
  DrawDowns.DrawNumber,
  Banks.BankshortName,
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.FAID,
  Units.UnitInvoiceNum,
  Vendors.VendorName,
  DrawDowns.DrawName,
  DrawDowns.DrawNumber,
  Banks.BankshortName;
