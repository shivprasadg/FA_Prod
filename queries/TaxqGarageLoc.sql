SELECT
  Locations.LState,
  Locations.LCounty,
  Locations.LCity,
  Locations.LAddrLine1,
  Locations.LZip,
  Clients.ClientShNm,
  Locations.LocationType,
  Locations.LocationName,
  Locations.LTaxRate,
  Count(Units.UnitID) AS CountOfUnitID
FROM
  (
    Clients
    INNER JOIN Locations ON Clients.clientID = Locations.ClientREF
  )
  INNER JOIN Units ON Locations.LocationID = Units.UnitGarageLocationRef
WHERE
  (
    (
      (Units.unitstatus)= "a"
      Or (Units.unitstatus)= "b"
      Or (Units.unitstatus)= "c"
      Or (Units.unitstatus)= "tr"
    )
  )
GROUP BY
  Locations.LState,
  Locations.LCounty,
  Locations.LCity,
  Locations.LAddrLine1,
  Locations.LZip,
  Clients.ClientShNm,
  Locations.LocationType,
  Locations.LocationName,
  Locations.LTaxRate,
  Clients.Clienttype,
  Locations.locobsolete
HAVING
  (
    (
      (Clients.ClientShNm) Not Like "demo*"
      And (Clients.ClientShNm)<> "test"
    )
    AND (
      (Clients.Clienttype)= "client"
      Or (Clients.Clienttype)= "guarantor"
    )
    AND (
      (Locations.locobsolete)= No
    )
  )
ORDER BY
  Locations.LState,
  Locations.LCounty,
  Locations.LCity,
  Locations.LZip;
