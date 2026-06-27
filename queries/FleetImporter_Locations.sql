SELECT
  DISTINCT Locations.LocationID,
  Locations.LocationType,
  Locations.LocationName,
  Locations.Locationsn,
  Locations.ClientREF,
  Locations.LocationType
FROM
  Locations
WHERE
  (
    (
      (Locations.ClientREF)= Forms!FleetStudy_Importer!ClientID
    )
    And (
      (Locations.LocationType)= "garage"
    )
  );
