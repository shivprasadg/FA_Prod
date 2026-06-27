SELECT
  Buyers.ClientShNm,
  RMKT.InvoiceRequestNum,
  RMKT.InvoiceRequestDate,
  RMKT.InvoiceSentDate,
  RMKT.InvoicePaidDate,
  RMKT.UnitRef,
  RMKT.BuyerRef
FROM
  RMKT
  INNER JOIN Buyers ON RMKT.BuyerRef = Buyers.clientID
WHERE
  (
    (
      (RMKT.InvoiceRequestNum) Is Not Null
    )
    AND (
      (RMKT.InvoicePaidDate) Is Null
    )
  )
ORDER BY
  Buyers.ClientShNm;
