SELECT
  FS.VIN,
  L.LocationID,
  L.LocationName
FROM
  FleetStudyImport_Work AS FS
  INNER JOIN Locations AS L ON (FS.Location = L.LocationName)
  AND (
    FS.ClientGroupID = L.ClientGroupID
  );
