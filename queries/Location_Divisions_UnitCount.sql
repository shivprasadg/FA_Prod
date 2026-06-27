SELECT
  Locations.ClientREF AS ClientID,
  Nz([Locations].[DivisionID], 0) AS DivisionID,
  Units.UnitGarageLocationRef AS GarageLocationID,
  Locations.LocationName,
  Locations.Locationsn,
  Count(Units.UnitID) AS Units
FROM
  Units
  INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
GROUP BY
  Locations.ClientREF,
  Nz([Locations].[DivisionID], 0),
  Units.UnitGarageLocationRef,
  Locations.LocationName,
  Locations.Locationsn;
