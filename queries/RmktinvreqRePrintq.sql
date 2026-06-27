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
  Clients.clientZip,
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
  Units.unitvin,
  [LCity] & " " & [LCounty] & " " & [LState] AS srt,
  lookup_States.StTaxType,
  Clients.clientZip,
  Clients.clientZip,
  RmktBills.Billdt,
  RmktBills.RBillDueDt,
  Rmkt.InvoiceRequestDate,
  RmktBills.RevNo,
  RmktBills.ID
FROM
  (
    (
      (
        (
          Units
          INNER JOIN Rmkt ON Units.UnitID = Rmkt.UnitRef
        )
        INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
      )
      INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    )
    INNER JOIN lookup_States ON Clients.clientState = lookup_States.StateAbbrev
  )
  INNER JOIN RmktBills ON Clients.clientID = RmktBills.BuyerRef
WHERE
  (
    (
      (Rmkt.BuyerRef)= forms!ReMarketInvoiceManager!buyerref
    )
    And (
      (Rmkt.InvoiceRequestDate)= forms!ReMarketInvoiceManager!RequestDate
    )
    And (
      (RmktBills.ID)= forms!ReMarketInvoiceManager!id
    )
  );
