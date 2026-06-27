SELECT
  Rmkt.UnitRef AS UnitID,
  Rmkt.InvoiceRequestNum AS Num,
  Format(
    [InvoiceRequestDate], "mm/dd/yy"
  ) AS Requested,
  Rmkt.RmkSm AS Rep,
  vw_SixKeys.AssetID,
  vw_SixKeys.Unitnum AS [Unit#],
  vw_SixKeys.VIN,
  [MYear] & " " & [Make] & " " & [Model] & IIf(
    Nz([SubType], "")= "",
    "",
    " " & [SubType]
  )& IIf([Type] = "Tractor", "", " " & [Type]) AS Description,
  vw_SixKeys.Status,
  Rmkt.SoldPrice AS [Sold Price],
  Rmkt.InvoiceSentDate,
  Rmkt.InvoicePaidDate AS Paid,
  vw_SixKeys.ClientShortName AS SourceCompany,
  [Schedule] & "|" & [UnitGroup] AS [Sch|Grp],
  Rmkt.BuyerRef,
  Rmkt.TargetSale
FROM
  (
    vw_SixKeys
    INNER JOIN Rmkt ON vw_SixKeys.UnitID = Rmkt.UnitRef
  )
  INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
WHERE
  (
    (
      (Rmkt.InvoiceRequestNum) Is Not Null
    )
  )
ORDER BY
  Rmkt.InvoiceRequestNum DESC,
  vw_SixKeys.AssetID,
  vw_SixKeys.Status;
