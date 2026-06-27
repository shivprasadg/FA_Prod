SELECT
  Locations.*,
  Locations.LocationName
FROM
  Locations
WHERE
  (
    (
      (Locations.LocObsolete)= No
    )
    AND (
      (Locations.LocationName) Like "*Remit*"
    )
  );
