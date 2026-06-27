SELECT
  Buyer.BuyerID,
  RmktInvoiceRequestTbl.InvoiceRequestID AS InvoiceRequestNum,
  RmktInvoiceRequestTbl.InvoiceRequestID,
  Units.UnitID AS AssetID,
  Units.unitunitnum,
  Units.UnitVIN,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  IIf(
    Nz([unitSubType], "")= "",
    "",
    [UnitSubType]
  )& IIf(
    [UnitType] = "Tractor", "", " " & [UnitType]
  ) AS Stype,
  Units.unittype,
  RmktInvoiceRequestTbl.InvoiceRequestDate,
  RmktInvoiceRequestTbl.InvoiceSentDate,
  Buyer.BuyerName,
  0 AS InvoiceRequested,
  Units.unitstatus,
  Rmkt.InvoicePaidDate,
  Buyer.BuyerShortName,
  Buyer.BuyerID,
  Rmkt.RmktDeposit,
  Rmkt.rmktdepdt,
  Rmkt.pbos,
  Rmkt.SoldPrice
FROM
  (
    (
      (
        Units
        INNER JOIN RMKT_BuyersOrders ON Units.UnitID = RMKT_BuyersOrders.UnitID
      )
      LEFT JOIN Rmkt ON RMKT_BuyersOrders.UnitID = Rmkt.UnitRef
    )
    INNER JOIN RmktInvoiceRequestTbl ON RMKT_BuyersOrders.BuyersOrderID = RmktInvoiceRequestTbl.InvoiceRequestID
  )
  INNER JOIN Buyer ON RmktInvoiceRequestTbl.BuyerID = Buyer.BuyerID;
