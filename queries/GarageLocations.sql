SELECT
  *
FROM
  Locations
WHERE
  (
    (
      (Locations.LocationType)= "Garage"
    )
  )
ORDER BY
  Locations.ClientREF,
  Locations.LocationName;
