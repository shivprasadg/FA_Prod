SELECT
  Locations.ClientREF,
  Locations.ClientGroupID,
  Locations.LocationID,
  Locations.DivisionID,
  Locations.LocationName,
  Locations.LState AS LState,
  IIf(
    Count([UnitID])> 0,
    "Y",
    ""
  ) AS Units,
  IIf(
    Count([ContactID])> 0,
    "Y"
  ) AS Contacts,
  [LocationName] & IIf(
    [LocationName] Like '*' & [LState] & '*',
    "", " " & [LState]
  ) AS Location,
  Locations.LocationType,
  Locations.LocObsolete,
  Format([LTaxRate], "Percent") AS TaxRate,
  Count(Units.unitId) AS UnitCount
FROM
  (
    Locations
    LEFT JOIN Units ON Locations.LocationID = Units.UnitGarageLocationRef
  )
  LEFT JOIN ContactsLinkedToLocation ON Locations.LocationID = ContactsLinkedToLocation.LocationID
GROUP BY
  Locations.ClientREF,
  Locations.ClientGroupID,
  Locations.LocationID,
  Locations.DivisionID,
  Locations.LocationName,
  Locations.LState,
  Locations.LocationType,
  Locations.LocObsolete,
  Format([LTaxRate], "Percent"),
  Locations.LCity
ORDER BY
  [LocationName] & IIf(
    [LocationName] Like '*' & [LState] & '*',
    "", " " & [LState]
  );
