SELECT
  Clients.ClientShNm,
  Locations.LocationName,
  Locations.LCity,
  Locations.Locationsn,
  Locations.LState,
  Clients.clientID,
  Locations.LocationID
FROM
  Clients
  LEFT JOIN Locations ON (
    Clients.clientID = Locations.ClientREF
  )
  AND (
    Clients.clientID = Locations.ClientREF
  )
WHERE
  (
    (
      (Locations.LCity) Like '*' & [LState] & '*'
    )
    AND (
      (Locations.LState)<> ""
      And (Locations.LState)<> "OK"
      And (Locations.LState)<> "HI"
    )
  )
  OR (
    (
      (Locations.LocationName) Like '*' & [LState] & '*'
    )
    AND (
      (Locations.LState)<> ""
      And (Locations.LState)<> "OK"
      And (Locations.LState)<> "HI"
    )
  )
ORDER BY
  Clients.ClientShNm,
  Locations.LocationName,
  Locations.LState;
