SELECT
  Units.unitunitnum,
  Units.UnitVIN,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  Rmkt.InvoiceRequested,
  Rmkt.InvoiceRequestDate,
  Rmkt.SoldPrice,
  Rmkt.InvoiceSentDate,
  Clients.clientCompanyName,
  Locations.LocationType,
  Locations.*
FROM
  (
    (
      Units
      INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
    )
    INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
  )
  INNER JOIN Locations ON Clients.clientID = Locations.ClientREF
WHERE
  (
    (
      (Rmkt.InvoiceRequestDate)= forms!ReMarketInvoiceManager!reqdt
    )
    And (
      (Locations.LocationType)= "billing"
    )
    And (
      (Rmkt.BuyerRef)= forms!ReMarketInvoiceManager!buyerref
    )
  );
