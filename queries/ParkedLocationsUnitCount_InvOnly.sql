SELECT
  Locations.ClientREF,
  Locations.LocationID,
  Locations.DivisionID,
  Locations.LocationName,
  Locations.LCity,
  Locations.LState AS LState,
  IIf(
    Sum([infleet])> 0,
    "Y",
    ""
  ) AS Units,
  IIf(
    Count([ContactID])> 0,
    Count([ContactID])
  ) AS Contacts,
  Locations.LocationType,
  Locations.LocObsolete
FROM
  (
    Locations
    INNER JOIN UnitsParkedLocations ON Locations.LocationID = UnitsParkedLocations.GarageLocationID
  )
  LEFT JOIN ContactsLinkedToLocationQuery ON Locations.LocationID = ContactsLinkedToLocationQuery.LocationID
GROUP BY
  Locations.ClientREF,
  Locations.LocationID,
  Locations.DivisionID,
  Locations.LocationName,
  Locations.LCity,
  Locations.LState,
  Locations.LocationType,
  Locations.LocObsolete
ORDER BY
  Locations.LocationName;
