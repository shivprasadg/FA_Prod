SELECT
  Locations.LocationID,
  Locations.LocationName,
  Locations.LocationStreet,
  Locations.LAddrLine1,
  Locations.LAddrLine2,
  Locations.LCity,
  Locations.LState,
  Locations.LZip
FROM
  Locations
WHERE
  (
    (
      (Locations.ClientREF)= forms!RemarketingClientView!clientid
    )
    And (
      (Locations.LocObsolete)= No
    )
  );
