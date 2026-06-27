SELECT
  Locations.LocationID,
  Locations.Locationsn,
  Locations.LocationStreet,
  Locations.LCity,
  Locations.LState,
  Locations.LZip
FROM
  Locations
WHERE
  (
    (
      (Locations.ClientREF)= Forms!UnitsForm!clientID
    )
    And (
      (Locations.LocationType)= "garage"
    )
    And (
      (Locations.LocObsolete)= No
    )
  )
ORDER BY
  Locations.LState,
  Locations.LCity,
  Locations.LocationName;
