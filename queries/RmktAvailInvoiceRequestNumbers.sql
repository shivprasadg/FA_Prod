SELECT
  Clients.ClientShNm AS Buyer,
  RmktInvoiceRequestTbl.InvoiceRequestID,
  RmktInvoiceRequestTbl.InvoiceRequestDate,
  Switch(
    Not IsDate([Rmkt].[InvoiceRequestDate]),
    "Open",
    IsDate([Rmkt].[InvoiceRequestDate])
    And Not IsDate([Rmkt].[InvoicePaidDate])
    And Not IsDate([Rmkt].[InvoiceSentDate]),
    "Request",
    Not IsDate([Rmkt].[InvoicePaidDate])
    And IsDate([Rmkt].[InvoiceSentDate]),
    "Sent",
    IsDate([Rmkt].[InvoicePaidDate]),
    "Paid"
  ) AS Status,
  Count(Rmkt.UnitRef) AS Units,
  RmktInvoiceRequestTbl.BuyerID,
  Rmkt.InvoiceSentDate,
  Rmkt.InvoicePaidDate,
  Rmkt.InvoiceRequestDate AS UnitRequestDate,
  RmktInvoiceRequestTbl.InvoiceSentDate AS Sent,
  RmktInvoiceRequestTbl.InvoicePaidDate AS Paid
FROM
  (
    RmktInvoiceRequestTbl
    LEFT JOIN Rmkt ON RmktInvoiceRequestTbl.InvoiceRequestID = Rmkt.InvoiceRequestNum
  )
  INNER JOIN Clients ON RmktInvoiceRequestTbl.BuyerID = Clients.clientID
GROUP BY
  Clients.ClientShNm,
  RmktInvoiceRequestTbl.InvoiceRequestID,
  RmktInvoiceRequestTbl.InvoiceRequestDate,
  RmktInvoiceRequestTbl.BuyerID,
  Rmkt.InvoiceSentDate,
  Rmkt.InvoicePaidDate,
  Rmkt.InvoiceRequestDate,
  RmktInvoiceRequestTbl.InvoiceSentDate,
  RmktInvoiceRequestTbl.InvoicePaidDate
ORDER BY
  RmktInvoiceRequestTbl.InvoiceRequestID DESC;
