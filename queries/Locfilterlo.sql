SELECT
  Locations.*
FROM
  Locations
WHERE
  (
    (
      (Locations.ClientREF)= [forms]![ClientsViewlo]![clientid]
    )
  );
