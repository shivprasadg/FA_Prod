SELECT
  Locations.*
FROM
  Locations
WHERE
  (
    (
      (Locations.LocationType)= "Billing"
    )
  );
