SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.AssetID,
  vw_SixKeys.VIN,
  vw_SixKeys.Unitnum,
  vw_SixKeys.Status,
  Nz([clientShNm], "none") AS Buyer,
  Rmkt.BuyOutDate,
  Rmkt.BuyOutCost,
  Rmkt.SoldPrice,
  Rmkt.InvoicePaidDate,
  vw_SixKeys.MYear,
  [Make] & " - " & [Model] & " - " & [type] AS [Desc],
  Buyers.clientID AS BuyerId,
  Rmkt.TargetSale,
  Rmkt.[SUF Sent],
  Rmkt.[SUF Returned],
  Rmkt.ParkDT,
  Units.unitoffleasedt
FROM
  (
    (
      vw_SixKeys
      INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
    )
    INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
  )
  LEFT JOIN Buyers ON Rmkt.BuyerRef = Buyers.clientID
ORDER BY
  vw_SixKeys.CompanyName,
  vw_SixKeys.FAID,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup;
