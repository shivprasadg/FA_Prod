SELECT
  RmktInvoiceRequestTbl.BuyerID,
  Buyer.BuyerShortName AS Buyer,
  RmktInvoiceRequestTbl.InvoiceRequestID AS Req,
  RmktInvoiceRequestTbl.InvoiceNumber AS InvNum,
  Format(
    [Rmkt].[InvoiceSentDate], "mm/dd/yy"
  ) AS Sent,
  Units.UnitID AS AssetID,
  Units.UnitunitNum AS UnitNum,
  Units.UnitVin AS Vin,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] AS UnitDesc,
  Format(
    [rmkt].[InvoicePaidDate], "mm/dd/yy"
  ) AS Paid
FROM
  (
    (
      RmktInvoiceRequestTbl
      INNER JOIN Rmkt ON RmktInvoiceRequestTbl.InvoiceRequestID = Rmkt.InvoiceRequestNum
    )
    INNER JOIN Units ON Rmkt.UnitRef = Units.UnitID
  )
  INNER JOIN Buyer ON RmktInvoiceRequestTbl.BuyerID = Buyer.BuyerID;
