DELETE DISTINCTROW Rmkt.*,
Rmkt_WrongStatus.Inspected,
Rmkt.BuyerRef,
Rmkt.BuyOutCost,
Rmkt.InvoicePaidDate,
Rmkt.[SUF Sent]
FROM
  Rmkt
  INNER JOIN Rmkt_WrongStatus ON Rmkt.UnitRef = Rmkt_WrongStatus.AssetID
WHERE
  (
    (
      (Rmkt_WrongStatus.Inspected) Is Null
    )
    AND (
      (Rmkt.BuyerRef) Is Null
    )
    AND (
      (Rmkt.BuyOutCost) Is Null
    )
    AND (
      (Rmkt.InvoicePaidDate) Is Null
    )
    AND (
      (Rmkt.[SUF Sent]) Is Null
    )
  );
