SELECT
  vw_SixKeys.AssetID,
  vw_SixKeys.Unitnum AS [Unit#],
  vw_SixKeys.VIN,
  [MYear] & " " & [Make] & " " & [Model] & IIf(
    Nz([SubType], "")= "",
    "",
    " " & [SubType]
  )& IIf([Type] = "Tractor", "", " " & [Type]) AS Description,
  vw_SixKeys.Status,
  Rmkt.TargetSale,
  vw_SixKeys.ClientShortName AS SourceCompany,
  [Schedule] & "|" & [UnitGroup] AS [Sch|Grp],
  RmktInvoiceRequestTbl.InvoiceRequestID,
  RmktInvoiceRequestTbl.InvoiceSentDate,
  RmktInvoiceRequestTbl.InvoicePaidDate,
  RmktInvoiceRequestTbl.InvoiceRequestDate,
  RmktInvoiceRequestTbl.BuyerID
FROM
  (
    (
      vw_SixKeys
      LEFT JOIN Rmkt ON vw_SixKeys.UnitID = Rmkt.UnitRef
    )
    LEFT JOIN RMKT_BuyersOrders ON vw_SixKeys.UnitID = RMKT_BuyersOrders.UnitID
  )
  LEFT JOIN RmktInvoiceRequestTbl ON RMKT_BuyersOrders.BuyersOrderID = RmktInvoiceRequestTbl.InvoiceRequestID
WHERE
  (
    (
      (vw_SixKeys.Status)= "R"
      Or (vw_SixKeys.Status)= "PS"
      Or (vw_SixKeys.Status)= "TR"
      Or (vw_SixKeys.Status)= "P"
    )
  )
ORDER BY
  vw_SixKeys.AssetID,
  vw_SixKeys.Status,
  vw_SixKeys.ClientShortName;
