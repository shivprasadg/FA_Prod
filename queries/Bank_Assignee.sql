SELECT
  Clients.clientID,
  Clients.ClientShNm AS ShortName,
  Clients.clientCompanyName AS BankName,
  Clients.clientType,
  Locations.LocationType,
  Locations.LocationStreet,
  [clientCompanyName] & Chr(13)& Chr(10)& [LAddrLine1] & Chr(13)& Chr(10)& IIf(
    IsNull([LAddrLine2]),
    Null,
    [LAddrLine2] & Chr(13)& Chr(10)
  )& IIf(
    IsNull([LAddrLine3]),
    Null,
    [LAddrLine3] & Chr(13)& Chr(10)
  )& [LCity] & " " & [LState] & " " & Left([LZip], 5)& IIf(
    Len([LZip])> 5,
    "-" & Right([LZip], 4),
    Null
  ) AS BankAddress,
  Locations.LAddrLine1,
  Locations.LAddrLine2,
  Locations.LAddrLine3,
  Locations.LCity,
  Locations.LCounty,
  Locations.LState,
  Locations.LZip,
  Locations.LocObsolete
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
      (Clients.clientType)= "bank"
    )
    AND (
      (Locations.LocationType)= "Billing"
    )
    AND (
      (Locations.LocObsolete)= False
    )
  )
ORDER BY
  Clients.ClientShNm;
