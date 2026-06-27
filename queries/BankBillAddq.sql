SELECT
  Clients.clientID,
  Locations.LocObsolete,
  Locations.*
FROM
  Clients
  INNER JOIN Locations ON Clients.clientID = Locations.ClientREF
WHERE
  (
    (
      (Locations.LocObsolete)= No
    )
    AND (
      (Clients.clienttype)= "bank"
    )
    AND (
      (Locations.LocationType)= "billing"
    )
  );
