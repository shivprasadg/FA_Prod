SELECT
  Locations.LocationID,
  Locations.LocationName AS [Loc Nm],
  Locations.LAddrLine1 AS [Str Add],
  Locations.LCity AS City,
  Locations.LState AS St,
  Locations.LZip AS Zip
FROM
  Locations
WHERE
  (
    (
      (Locations.LocationType)= "Garage"
    )
    AND (
      (Locations.ClientREF) Is Not Null
    )
    AND (
      (Locations.LocObsolete)= No
    )
  )
ORDER BY
  Locations.LState,
  Locations.LCity,
  Locations.LocationName;
