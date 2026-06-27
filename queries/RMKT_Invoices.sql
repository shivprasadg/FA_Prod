SELECT
  RmktInvoiceRequestTbl.InvoiceRequestID,
  RmktInvoiceRequestTbl.InvoiceRequestDate,
  RmktInvoiceRequestTbl.InvoicePaidDate,
  RmktInvoiceRequestTbl.BuyerID,
  Switch(
    Not IsDate(
      [RmktInvoiceRequestTbl].[InvoiceRequestDate]
    ),
    "Open",
    IsDate(
      [RmktInvoiceRequestTbl].[InvoiceRequestDate]
    )
    And Not IsDate(
      [RmktInvoiceRequestTbl].[InvoicePaidDate]
    )
    And Not IsDate(
      [RmktInvoiceRequestTbl].[InvoiceSentDate]
    ),
    "Request",
    Not IsDate(
      [RmktInvoiceRequestTbl].[InvoicePaidDate]
    )
    And IsDate(
      [RmktInvoiceRequestTbl].[InvoiceSentDate]
    ),
    "Sent",
    IsDate(
      [RmktInvoiceRequestTbl].[InvoicePaidDate]
    ),
    "Paid"
  ) AS Status,
  RmktInvoiceRequestTbl.InvoiceNumber,
  RmktInvoiceRequestTbl.InvoiceDate,
  RmktInvoiceRequestTbl.InvoiceDueDate,
  RmktInvoiceRequestTbl.InvoiceSentDate,
  RmktInvoiceRequestTbl.InvoiceLink,
  RmktInvoiceRequestTbl.TaxRateToUse
FROM
  RmktInvoiceRequestTbl
WHERE
  (
    (
      (
        RmktInvoiceRequestTbl.InvoiceRequestID
      ) Is Not Null
    )
  );
