UPDATE
  (
    Units
    INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
  )
  INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
SET
  Rmkt.RmkSm = [forms]![Purchaser Invoice Request]![RSM],
  Rmkt.InvoiceRequestDate = Date()
WHERE
  (
    (
      (Clients.clientID)= [forms]![RemarketingClientView]![clientid]
    )
    AND (
      (Rmkt.InvoicePaidDate) Is Null
    )
    AND (
      (Rmkt.InvoiceRequested)= Yes
    )
    AND (
      (Units.unitstatus)= "tr"
      Or (Units.unitstatus)= "r"
      Or (Units.unitstatus)= "k"
    )
  );
