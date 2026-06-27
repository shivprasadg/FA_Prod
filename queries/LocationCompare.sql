SELECT
  Units.ClientID,
  Units.UnitID,
  Units.UnitVIN,
  Locations.LocationName,
  Locations_1.LocationName AS REALClientLoc,
  Units.UnitGarageLocationRef,
  Locations_1.LocationID AS RealLocID,
  Clients.ClientShNm AS ListedClient,
  Clients_1.ClientShNm AS RealClient,
  Clients.clientID AS ListedClientID,
  Locations_1.ClientREF AS RealClientID
FROM
  (
    (
      Units
      INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
    )
    INNER JOIN Clients ON Locations.ClientREF = Clients.clientID
  )
  INNER JOIN (
    Locations AS Locations_1
    INNER JOIN Clients AS Clients_1 ON Locations_1.ClientREF = Clients_1.clientID
  ) ON Locations.LocationName = Locations_1.LocationName
WHERE
  (
    (
      (Units.ClientID)= [AskClientID]
    )
    AND (
      (Locations_1.ClientREF)= [AskClientID]
    )
  );
