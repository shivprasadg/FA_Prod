SELECT
  Locations.LocationID,
  Locations.LState,
  Locations.LCity,
  Locations.LZip,
  TaxCounties.CountyName,
  Locations.LCounty,
  Locations.LCountyID,
  Locations.LStateID,
  Locations.LocationType,
  DLookUp(
    "County",
    "TaxZipFipsMatrix",
    "Zip='" & Left([LZip], 5)& "'"
  ) AS LookCounty,
  DLookUp(
    "ID", "TaxCounties", "TaxStatesID=" & [AskStateId] & "  AND CountyName='" & [LookCounty] & "'"
  ) AS LookCountyID INTO LocationsCountyTemp
FROM
  TaxCounties
  INNER JOIN Locations ON (
    TaxCounties.CountyName = Locations.LCounty
  )
  AND (
    TaxCounties.TaxStatesID = Locations.LStateID
  )
WHERE
  (
    (
      (Locations.LState)= [AskState]
    )
    AND (
      (Locations.LocationType)<> "Title"
    )
  );
