SELECT
  RBO.BuyersOrderID,
  RBO.UnitID,
  U.UnitUnitNum,
  U.UnitVIN,
  tReq.InvoiceDate,
  tReq.InvoiceSentDate,
  tReq.InvoicePaidDate,
  R.TargetSale,
  DateDiff(
    "d",
    tReq.InvoiceSentDate,
    Nz(
      tReq.InvoicePaidDate,
      Date()
    )
  ) AS Days,
  RBO.SalePrice,
  tReq.FundingType,
  CStr(UnitYr)& "  " & U.UnitMake & "  " & U.UnitModel & "  " & U.UnitSubType & "  " & U.UnitType + ' ' + U.Axle AS UnitDesc,
  tReq.SP,
  U.UnitStatus AS STS,
  tReq.BuyerID,
  RBO.WarrantyPrice
FROM
  (
    (
      RMKT_BuyersOrders AS RBO
      INNER JOIN RmktInvoiceRequestTbl AS tReq ON RBO.BuyersOrderID = tReq.InvoiceRequestID
    )
    INNER JOIN Units AS U ON RBO.UnitID = U.UnitID
  )
  INNER JOIN Rmkt AS R ON U.UnitID = R.UnitRef;
