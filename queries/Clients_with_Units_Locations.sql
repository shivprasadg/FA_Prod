SELECT
  Clients_with_Units_Address.ClientGroupId,
  Clients_with_Units_Address.clientID,
  Nz(
    [ClientDivisions].[DivisionID],
    0
  ) AS DivisionID,
  Locations.LocationID,
  Clients_with_Units_Address.clientCompanyName,
  Nz(
    [ClientDivisions].[DivisionName],
    "All"
  ) AS DivisionName,
  Locations.LocationType,
  Locations.LocationName,
  Location_Divisions_UnitCount.Locationsn,
  Location_Divisions_UnitCount.Units AS [Units@Location],
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LAddrLine2,
  Locations.LCity,
  Locations.LState,
  Locations.LZip
FROM
  (
    (
      Clients_with_Units_Address
      INNER JOIN Locations ON Clients_with_Units_Address.clientID = Locations.ClientREF
    )
    LEFT JOIN ClientDivisions ON Locations.DivisionID = ClientDivisions.DivisionID
  )
  INNER JOIN Location_Divisions_UnitCount ON Locations.LocationID = Location_Divisions_UnitCount.GarageLocationID;
