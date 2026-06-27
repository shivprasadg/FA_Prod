SELECT
  Units.UnitID AS AssetID,
  Rmkt.InvoicePaidDate,
  Units.unitstatus,
  vw_SixKeys.CompanyName,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.VIN,
  vw_SixKeys.FAID,
  Buyers.ClientShNm AS Buyer,
  Rmkt.SoldPrice
FROM
  (
    Rmkt
    INNER JOIN Buyers ON Rmkt.BuyerRef = Buyers.clientID
  )
  INNER JOIN (
    Units
    LEFT JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.AssetID
  ) ON Rmkt.UnitRef = Units.UnitID
WHERE
  (
    (
      (Rmkt.InvoicePaidDate) Is Not Null
    )
    AND (
      (Units.unitstatus)= "S"
      Or (Units.unitstatus)= "D"
      Or (Units.unitstatus)= "KS"
      Or (Units.unitstatus)= "PS"
    )
  )
ORDER BY
  Rmkt.InvoicePaidDate DESC;
