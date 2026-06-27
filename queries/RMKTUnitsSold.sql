SELECT
  Rmkt.RmkSm,
  Rmkt.UnitRef AS AssetID,
  Rmkt.InvoicePaidDate,
  Rmkt.SoldPrice,
  RMKTSalesPersons.ID AS SpID
FROM
  Rmkt
  INNER JOIN RMKTSalesPersons ON Rmkt.RmkSm = RMKTSalesPersons.SP
WHERE
  (
    (
      (Rmkt.InvoicePaidDate) Is Not Null
    )
  );
