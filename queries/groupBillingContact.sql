SELECT
  Locations.LocationID,
  Locations.LocationName,
  Locations.Locationsn,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LCity,
  Locations.LState,
  Locations.LZip,
  Locations.LPhone,
  Locations.LFax,
  Locations.LocObsdate,
  Locations.LocObsCmt
FROM
  Locations
WHERE
  (
    (
      (Locations.ClientREF)= Forms!Groups_view!clientRef
    )
    And (
      (Locations.LocationType)= "billing"
    )
  );
