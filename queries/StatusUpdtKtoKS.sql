UPDATE
  (
    (
      Units
      INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
    )
    INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
  )
  LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
SET
  Units.unitstatus = "KS"
WHERE
  (
    (
      (Units.unitstatus)= "K"
    )
    And (
      (Rmkt.BuyerRef)= forms!ReMarketInvoiceManager!buyerref
    )
    And (
      (Rmkt.InvoicePaidDate) Is Not Null
    )
    And (
      (Rmkt.InvoiceRequestDate)= forms!ReMarketInvoiceManager!reqdt
    )
  );
