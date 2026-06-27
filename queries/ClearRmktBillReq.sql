UPDATE
  (
    Units
    INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
  )
  INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
SET
  Rmkt.InvoiceRequested = 0
WHERE
  (
    (
      (Rmkt.InvoiceRequested)= True
    )
    And (
      (Clients.clientID)= Forms!RemarketingClientView!clientid
    )
  );
