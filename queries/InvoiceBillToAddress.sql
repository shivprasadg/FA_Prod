SELECT
  Locations.LocationID AS BillingToID,
  [LocationName] & " " & [LCity] & ", " & [LState] AS Location,
  [LocationStreet] & Chr(13)& Chr(10)& IIf(
    IsNull([LAddrLine1]),
    Null,
    [Locations].[LAddrLine1] & Chr(13)& Chr(10)
  )& IIf(
    IsNull([Locations].[LAddrLine2]),
    Null,
    [Locations].[LAddrLine2] & Chr(13)& Chr(10)
  )& [Locations].[LCity] & " " & [Locations].[LState] & " " & [Locations].[LZip] AS BillingAddress
FROM
  Locations
WHERE
  (
    (
      (Locations.LocObsolete)= No
    )
  );
