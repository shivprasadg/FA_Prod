SELECT
  Units.UnitID,
  Rmkt.RmktId,
  Rmkt.BuyerRef,
  Clients.clientID,
  Clients.clientCompanyName,
  Clients.clientCoNm2,
  Clients.clientDBA,
  Clients.clientAddress1,
  Clients.clientAddress2,
  Clients.clientCity,
  Clients.clientZip,
  Clients.clientZip AS Buyerzip,
  Clients.[Dlr/Resale],
  Units.unitvin,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype,
  Rmkt.SoldPrice,
  Clients.clientState,
  Rmkt.TargetSale,
  Rmkt.InvoiceRequested,
  Units.unitstatus,
  Rmkt.RmkSm,
  Rmkt.InvoicePaidDate,
  Rmkt.InvoiceRequestDate,
  Locations.*,
  Rmkt.pbos
FROM
  (
    (
      Units
      INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
    )
    INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
  )
  LEFT JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
WHERE
  (
    (
      (Rmkt.BuyerRef)= forms!ReMarketInvoiceManager!buyerref
    )
    And (
      (Rmkt.InvoicePaidDate) Is Not Null
    )
    And (
      (Rmkt.InvoiceRequestDate)= forms!ReMarketInvoiceManager!RequestDate
    )
    And (
      (Rmkt.pbos)= True
    )
  );
