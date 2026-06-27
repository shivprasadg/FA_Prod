SELECT
  Clients.clientCompanyName,
  Clients.clientAddress1,
  Clients.clientAddress2,
  Clients.clientCity,
  Clients.clientState,
  Clients.clientZip,
  Clients.[Dlr/Resale],
  Units.unitvin,
  Units.unityr,
  Units.unitmake,
  Units.unittype,
  Clients.clientID
FROM
  Units
  INNER JOIN (
    Rmkt
    INNER JOIN Clients ON Rmkt.BuyerRef = Clients.clientID
  ) ON Units.UnitID = Rmkt.UnitRef;
